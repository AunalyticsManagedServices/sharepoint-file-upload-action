# -*- coding: utf-8 -*-
"""
Parallel file upload orchestration for SharePoint sync.

This module provides parallel upload capabilities while maintaining 100%
compatibility with existing code, console output, and statistics tracking.
"""

import os
import time
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from .thread_utils import (
    ThreadSafeStatsWrapper,
    ThreadSafeSet,
    BatchQueue,
    enable_thread_safe_print
)
from .uploader import upload_file_with_structure, upload_file
from .markdown_converter import convert_markdown_to_html, rewrite_markdown_links
from .file_handler import sanitize_path_components
from .utils import is_debug_enabled
from .monitoring import rate_monitor


class ParallelUploader:
    """
    Parallel file upload orchestrator.

    Uploads multiple files concurrently while:
    - Maintaining exact same console output format
    - Preserving all statistics tracking
    - Respecting Graph API rate limits
    - Handling errors per-file like sequential mode
    """

    def __init__(self, max_workers=4, upload_stats_instance=None, batch_metadata_updates=True):
        """
        Initialize parallel uploader.

        Args:
            max_workers (int): Maximum concurrent upload threads (default: 4)
            upload_stats_instance: Reference to global upload_stats instance
            batch_metadata_updates (bool): Use batch updates for FileHash metadata
        """
        self.max_workers = max_workers
        self.batch_metadata = batch_metadata_updates

        # Wrap existing stats with thread-safety
        if upload_stats_instance:
            self.stats_wrapper = ThreadSafeStatsWrapper(upload_stats_instance.stats)
        else:
            # Fallback for testing
            self.stats_wrapper = ThreadSafeStatsWrapper({
                'new_files': 0,
                'replaced_files': 0,
                'skipped_files': 0,
                'failed_files': 0,
                'bytes_uploaded': 0,
                'bytes_skipped': 0,
                'compared_by_hash': 0,
                'compared_by_size': 0,
                'hash_new_saved': 0,
                'hash_updated': 0,
                'hash_matched': 0,
                'hash_save_failed': 0
            })

        # Thread-safe set for converted markdown files
        self.converted_md_files = ThreadSafeSet()

        # Queue for batch metadata updates
        self.metadata_queue = BatchQueue(batch_size=20) if self.batch_metadata else None

    def process_files(self, local_files, site_id, drive_id, root_item_id, base_path, config,
                     filehash_available, library_name, converted_md_files_set=None):
        """
        Process and upload files in parallel.

        Args:
            local_files (list): List of local file paths to process
            site_id (str): SharePoint site ID
            drive_id (str): SharePoint drive ID
            root_item_id (str): Root folder item ID
            base_path (str): Base path for folder structure
            config: Configuration object
            filehash_available (bool): Whether FileHash column exists
            library_name (str): SharePoint library name
            converted_md_files_set (set): Set to track converted markdown files

        Returns:
            int: Number of failed uploads
        """
        # Separate markdown files from regular files
        md_files = []
        regular_files = []

        for f in local_files:
            if os.path.isfile(f):
                if f.lower().endswith('.md') and config.convert_md_to_html:
                    md_files.append(f)
                else:
                    regular_files.append(f)

        failed_count = 0

        # Process markdown files first (may need conversion)
        if md_files:
            if is_debug_enabled():
                print(f"[MD] Processing {len(md_files)} markdown files...")

            failed_count += self._process_markdown_files_parallel(
                md_files, site_id, drive_id, root_item_id, base_path, config,
                filehash_available, library_name
            )

        # Process regular files in parallel
        if regular_files:
            if is_debug_enabled():
                print(f"[→] Uploading {len(regular_files)} files in parallel (workers: {self.max_workers})...")

            failed_count += self._upload_files_parallel(
                regular_files, site_id, drive_id, root_item_id, base_path, config,
                filehash_available, library_name
            )

        # Process any remaining batch metadata updates
        if self.metadata_queue:
            self._flush_metadata_queue(config, library_name)

        # Copy converted files back to provided set if given
        if converted_md_files_set is not None:
            for file in self.converted_md_files.copy():
                converted_md_files_set.add(file)

        return failed_count

    def _preprocess_markdown_file(self, file_path, base_path, config):
        """
        Preprocess a raw markdown file to rewrite internal links to SharePoint URLs.

        Creates a temporary markdown file with rewritten links for upload.
        This is used for .md files that are NOT being converted to HTML.

        Args:
            file_path (str): Path to the original markdown file
            base_path (str): Base path for relative path calculation
            config: Configuration object with SharePoint settings

        Returns:
            str: Path to temporary preprocessed markdown file, or original path if preprocessing fails
        """
        try:
            # Read original markdown content
            with open(file_path, 'r', encoding='utf-8') as f:
                md_content = f.read()

            # Calculate relative path for link rewriting
            if base_path:
                rel_path_str = os.path.relpath(file_path, base_path)
            else:
                rel_path_str = file_path

            # Normalize path separators
            rel_path_str = rel_path_str.replace('\\', '/')

            # Construct SharePoint base URL
            sharepoint_base_url = f"https://{config.sharepoint_host_name}/sites/{config.site_name}/Shared%20Documents/{config.upload_path}"

            # Rewrite internal links
            rewritten_content = rewrite_markdown_links(md_content, sharepoint_base_url, rel_path_str)

            # Check if any changes were made
            if rewritten_content == md_content:
                # No links were rewritten, use original file
                if is_debug_enabled():
                    print(f"[MD] No links to rewrite in: {file_path}")
                return file_path

            # Create temporary file with rewritten content
            temp_fd, temp_path = tempfile.mkstemp(suffix='.md', prefix='rewritten_md_')
            try:
                with os.fdopen(temp_fd, 'w', encoding='utf-8') as f:
                    f.write(rewritten_content)
            except Exception as write_error:
                os.close(temp_fd)
                raise write_error

            if is_debug_enabled():
                print(f"[MD] Preprocessed markdown with rewritten links: {file_path}")

            return temp_path

        except Exception as e:
            print(f"[!] Failed to preprocess markdown file {file_path}: {e}")
            # Fall back to original file
            return file_path

    def _upload_files_parallel(self, file_list, site_id, drive_id, root_item_id, base_path, config,
                               filehash_available, library_name):
        """
        Upload regular files in parallel.

        Returns:
            int: Number of failed uploads
        """
        failed_count = 0
        temp_files_to_cleanup = []  # Track temp files for cleanup

        def upload_worker(worker_id, filepath):
            """Worker function for parallel upload"""
            import threading

            # Name this thread for debug logging
            threading.current_thread().name = f"Upload-{worker_id}"

            # Enable thread-safe print for this thread
            enable_thread_safe_print()

            file_to_upload = filepath
            is_temp = False

            try:
                # Preprocess raw markdown files to rewrite links
                if filepath.lower().endswith('.md'):
                    preprocessed_path = self._preprocess_markdown_file(filepath, base_path, config)
                    if preprocessed_path != filepath:
                        # A temp file was created
                        file_to_upload = preprocessed_path
                        is_temp = True
                        temp_files_to_cleanup.append(preprocessed_path)

                # Call existing upload function - maintains all output/statistics
                upload_file_with_structure(
                    site_id, drive_id, root_item_id, file_to_upload, base_path,
                    config.tenant_url, library_name,
                    4*1024*1024,  # 4MB chunk size
                    config.force_upload,
                    filehash_available,
                    config.tenant_id, config.client_id, config.client_secret,
                    config.login_endpoint, config.graph_endpoint,
                    self.stats_wrapper,  # Thread-safe wrapper
                    config.max_retry,
                    metadata_queue=self.metadata_queue  # Pass queue for batch updates
                )
                return True

            except Exception as upload_err:
                # Error already logged by upload function
                print(f"[!] Upload failed for {filepath}: {str(upload_err)[:200]}")
                self.stats_wrapper.increment('failed_files')
                return False
            finally:
                # Clean up temp file if one was created
                if is_temp and os.path.exists(file_to_upload):
                    try:
                        os.remove(file_to_upload)
                    except Exception:
                        pass  # Ignore cleanup errors

        # Execute uploads in parallel
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Submit all upload tasks with worker IDs
            future_to_file = {
                executor.submit(upload_worker, idx % self.max_workers + 1, f): f
                for idx, f in enumerate(file_list)
            }

            # Process completed uploads
            for future in as_completed(future_to_file):
                file_path = future_to_file[future]
                try:
                    success = future.result()
                    if not success:
                        failed_count += 1

                    # Check if we should slow down due to rate limiting
                    if rate_monitor.should_slow_down():
                        time.sleep(1)  # Brief pause if approaching limits

                except Exception as e:
                    print(f"[!] Unexpected error processing {file_path}: {e}")
                    failed_count += 1
                    self.stats_wrapper.increment('failed_files')

        return failed_count

    def _process_markdown_files_parallel(self, md_files, site_id, drive_id, root_item_id, base_path,
                                        config, filehash_available, library_name):
        """
        Process markdown files in parallel (conversion + upload).

        Returns:
            int: Number of failed conversions/uploads
        """
        failed_count = 0

        def process_md_worker(worker_id, md_filepath):
            """Worker for markdown processing"""
            import threading

            # Name this thread for debug logging
            threading.current_thread().name = f"Convert-{worker_id}"

            enable_thread_safe_print()

            try:
                md_success = self._process_single_markdown_file(
                    md_filepath, site_id, drive_id, root_item_id, base_path, config,
                    filehash_available, library_name
                )
                if md_success:
                    self.converted_md_files.add(md_filepath)
                return md_success

            except Exception as md_err:
                print(f"[!] Markdown processing failed for {md_filepath}: {md_err}")
                return False

        # Process markdown files in parallel
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            future_to_file = {
                executor.submit(process_md_worker, idx % self.max_workers + 1, f): f
                for idx, f in enumerate(md_files)
            }

            for future in as_completed(future_to_file):
                md_file = future_to_file[future]
                try:
                    success = future.result()
                    if not success:
                        failed_count += 1
                except Exception as e:
                    print(f"[!] Unexpected error with {md_file}: {e}")
                    failed_count += 1

        return failed_count

    def _process_single_markdown_file(self, file_path, site_id, drive_id, root_item_id, base_path,
                                      config, filehash_available, library_name):
        """
        Process single markdown file (convert + upload).
        Mirrors logic from main.py:process_markdown_file()

        Returns:
            bool: True if successful
        """
        if is_debug_enabled():
            print(f"[MD] Converting markdown file: {file_path}")

        try:
            # Calculate hash of source .md file BEFORE conversion
            # This hash will be used for the converted .html file in SharePoint
            from .file_handler import calculate_file_hash
            md_file_hash = calculate_file_hash(file_path)
            if md_file_hash and is_debug_enabled():
                print(f"[#] Source .md file hash: {md_file_hash[:8]}... (will be used for .html file)")

            # Read markdown content
            with open(file_path, 'r', encoding='utf-8') as md_file_handle:
                md_content = md_file_handle.read()

            # Calculate relative path for SharePoint link rewriting
            if base_path:
                rel_path_str = os.path.relpath(file_path, base_path)
            else:
                rel_path_str = file_path

            # Normalize path separators to forward slashes
            rel_path_str = rel_path_str.replace('\\', '/')

            # Construct SharePoint base URL for link rewriting
            # Format: https://host/sites/sitename/Shared Documents/upload_path
            sharepoint_base_url = f"https://{config.sharepoint_host_name}/sites/{config.site_name}/Shared%20Documents/{config.upload_path}"

            # Convert to HTML with link rewriting
            html_content = convert_markdown_to_html(
                md_content,
                file_path,
                sharepoint_base_url=sharepoint_base_url,
                current_file_rel_path=rel_path_str
            )

            # Create temp HTML file
            temp_html_fd, html_path = tempfile.mkstemp(suffix='.html', prefix='converted_md_')

            try:
                with os.fdopen(temp_html_fd, 'w', encoding='utf-8') as html_file:
                    html_file.write(html_content)
            except Exception as write_error:
                os.close(temp_html_fd)
                raise write_error

            # Calculate paths
            original_html_path = file_path.replace('.md', '.html')

            # Get relative path
            if base_path:
                rel_path_str = os.path.relpath(original_html_path, base_path)
            else:
                rel_path_str = original_html_path

            # Normalize and sanitize (ensure str type)
            if isinstance(rel_path_str, bytes):
                rel_path_str = rel_path_str.decode('utf-8')
            rel_path_str = rel_path_str.replace('\\', '/')
            sanitized_rel_path = sanitize_path_components(rel_path_str)
            dir_path = os.path.dirname(sanitized_rel_path)

            # Determine target folder
            if dir_path and dir_path != "." and dir_path != "":
                from .uploader import ensure_folder_exists
                target_folder_id = ensure_folder_exists(
                    site_id, drive_id, root_item_id, dir_path,
                    config.tenant_id, config.client_id, config.client_secret,
                    config.login_endpoint, config.graph_endpoint
                )
            else:
                target_folder_id = root_item_id

            desired_html_filename = os.path.basename(original_html_path)

            # Upload HTML file with source .md file hash
            # This allows hash-based comparison instead of size-only (solves Mermaid SVG ID variation issue)
            for i in range(config.max_retry):
                try:
                    upload_file(
                        site_id, drive_id, target_folder_id, html_path, 4*1024*1024, config.force_upload,
                        config.tenant_url, library_name, filehash_available,
                        config.tenant_id, config.client_id, config.client_secret,
                        config.login_endpoint, config.graph_endpoint,
                        self.stats_wrapper, desired_name=desired_html_filename,
                        metadata_queue=self.metadata_queue,  # Pass queue for batch updates
                        pre_calculated_hash=md_file_hash,  # Use source .md file hash for comparison
                        display_path=sanitized_rel_path  # Show full relative path in debug output
                    )
                    break
                except Exception as e:
                    if i == config.max_retry - 1:
                        print(f"[Error] Failed to upload {original_html_path} after {config.max_retry} attempts")
                        raise e
                    else:
                        print(f"[!] Retrying upload... ({i+1}/{config.max_retry})")
                        time.sleep(2)

            # Clean up temp file
            if os.path.exists(html_path):
                os.remove(html_path)

            return True

        except Exception as e:
            print(f"[Error] Failed to convert markdown file {file_path}: {e}")
            # Fall back to uploading raw markdown
            try:
                upload_file_with_structure(
                    site_id, drive_id, root_item_id, file_path, base_path, config.tenant_url, library_name,
                    4*1024*1024, config.force_upload, filehash_available,
                    config.tenant_id, config.client_id, config.client_secret,
                    config.login_endpoint, config.graph_endpoint,
                    self.stats_wrapper, config.max_retry,
                    metadata_queue=self.metadata_queue  # Pass queue for batch updates
                )
                return True
            except Exception as fallback_error:
                print(f"[Error] Fallback markdown upload failed: {fallback_error}")
                return False

    def _flush_metadata_queue(self, config, library_name):
        """
        Flush any remaining metadata updates from queue.

        Args:
            config: Configuration object
            library_name (str): SharePoint library name
        """
        if not self.metadata_queue or self.metadata_queue.empty():
            return

        print(f"[#] Processing remaining metadata updates...")

        # Get all remaining items
        remaining = self.metadata_queue.get_all_remaining()
        if remaining:
            # Add delay for complex file types to allow SharePoint processing to complete
            # Different file types need processing time: virus scan, content indexing, conversion, sanitization
            html_count = sum(1 for _, filename, _, _, _ in remaining
                           if filename.lower().endswith('.html'))
            pdf_count = sum(1 for _, filename, _, _, _ in remaining
                          if filename.lower().endswith('.pdf'))
            office_count = sum(1 for _, filename, _, _, _ in remaining
                              if any(filename.lower().endswith(ext) for ext in ['.docx', '.xlsx', '.pptx', '.doc', '.xls', '.ppt']))
            image_count = sum(1 for _, filename, _, _, _ in remaining
                             if any(filename.lower().endswith(ext) for ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp', '.tiff']))
            complex_count = html_count + pdf_count + office_count + image_count
            total_count = len(remaining)

            if is_debug_enabled():
                simple_count = total_count - complex_count
                print(f"[DEBUG] Queue contains {total_count} items: {html_count} HTML, {pdf_count} PDF, {office_count} Office, {image_count} images, {simple_count} other")

            if complex_count > 0:
                import time
                # Delay based on file complexity
                if html_count > 0:
                    delay_seconds = 10  # HTML needs sanitization
                elif pdf_count > 0 or office_count > 0:
                    delay_seconds = 8  # PDFs and Office docs need processing
                else:
                    delay_seconds = 5  # Other files need basic processing

                if is_debug_enabled():
                    print(f"[⏱] Waiting {delay_seconds} seconds for SharePoint to process {complex_count} complex files...")
                time.sleep(delay_seconds)

            self._process_metadata_batch(remaining, config, library_name)

    def _process_metadata_batch(self, batch, config, library_name):
        """
        Process batch of metadata updates.

        Args:
            batch (list): List of (item_id, filename, hash_value, is_file_update, display_path) tuples
            config: Configuration object
            library_name (str): SharePoint library name
        """
        # Import here to avoid circular dependency
        from .graph_api import batch_update_filehash_fields

        if not batch:
            return

        if is_debug_enabled():
            print(f"[#] Batch updating {len(batch)} FileHash values...")

        # Extract update type info before sending to batch API
        # Map item_id to is_file_update flag
        update_types = {item_id: is_update for item_id, _, _, is_update, _ in batch}

        # Convert batch to format expected by batch_update_filehash_fields
        # It expects (item_id, filename, hash_value, display_path) tuples
        api_batch = [(item_id, filename, hash_value, display_path)
                     for item_id, filename, hash_value, _, display_path in batch]

        try:
            results = batch_update_filehash_fields(
                config.tenant_url, library_name, api_batch,
                config.tenant_id, config.client_id, config.client_secret,
                config.login_endpoint, config.graph_endpoint
            )

            # Collect ALL failed items for potential retry (not just HTML)
            failed_items = []

            # Create a lookup map for the original batch items
            # This ensures we can find all items efficiently
            batch_lookup = {}
            for orig_item_id, filename, hash_value, is_update, display_path in batch:
                # Convert to string to ensure consistent comparison
                batch_lookup[str(orig_item_id)] = (filename, hash_value, is_update, display_path)

            # Update statistics based on results and update type
            for item_id, success in results.items():
                # Convert to string for consistent comparison
                item_id_str = str(item_id)

                if success:
                    # Track based on whether this was new file or update
                    is_update = update_types.get(item_id, False)
                    if is_update:
                        self.stats_wrapper.increment('hash_updated')
                    else:
                        self.stats_wrapper.increment('hash_new_saved')
                else:
                    # Collect ALL failed items for retry (not just HTML)
                    if item_id_str in batch_lookup:
                        filename, hash_value, is_update, display_path = batch_lookup[item_id_str]
                        failed_items.append((item_id, filename, hash_value, is_update, display_path))
                    else:
                        # Item not found in batch_lookup - this shouldn't happen but log it
                        if is_debug_enabled():
                            print(f"[DEBUG] Warning: Failed item_id {item_id} not found in batch lookup")

                    self.stats_wrapper.increment('hash_save_failed')

            # Categorize failed items by file type for appropriate retry delays
            if failed_items:
                html_count = sum(1 for _, f, _, _, _ in failed_items if f.lower().endswith('.html'))
                pdf_count = sum(1 for _, f, _, _, _ in failed_items if f.lower().endswith('.pdf'))
                office_count = sum(1 for _, f, _, _, _ in failed_items if any(f.lower().endswith(ext) for ext in ['.docx', '.xlsx', '.pptx', '.doc', '.xls', '.ppt']))
                image_count = sum(1 for _, f, _, _, _ in failed_items if any(f.lower().endswith(ext) for ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp', '.tiff']))
                other_count = len(failed_items) - html_count - pdf_count - office_count - image_count

                if is_debug_enabled():
                    print(f"[DEBUG] Failed items by type: {html_count} HTML, {pdf_count} PDF, {office_count} Office, {image_count} images, {other_count} other")

            # Retry ALL failed files after additional delay
            # Different file types may need processing time (HTML sanitization, PDF scanning, Office conversion)
            if failed_items:
                import time
                # Determine retry delay based on file types
                # Different files need different processing time in SharePoint
                if html_count > 0 or office_count > 0:
                    retry_delay = 15  # Longer delay for files needing conversion/sanitization
                elif pdf_count > 0 or image_count > 0:
                    retry_delay = 12  # Medium delay for files needing scanning/thumbnails
                else:
                    retry_delay = 8  # Shorter delay for simpler files (text, scripts, etc.)

                if is_debug_enabled():
                    print(f"[⏱] {len(failed_items)} files failed (likely still processing).")
                    print(f"    Waiting {retry_delay} seconds before retry...")
                else:
                    print(f"[⏱] {len(failed_items)} files need retry after processing delay...")
                time.sleep(retry_delay)

                print(f"[#] Retrying {len(failed_items)} failed FileHash updates...")

                # Prepare retry batch (with display_path for better debug output)
                retry_api_batch = [(item_id, filename, hash_value, display_path)
                                  for item_id, filename, hash_value, _, display_path in failed_items]

                try:
                    retry_results = batch_update_filehash_fields(
                        config.tenant_url, library_name, retry_api_batch,
                        config.tenant_id, config.client_id, config.client_secret,
                        config.login_endpoint, config.graph_endpoint, batch_size=10  # Smaller batch size for retries
                    )

                    # Update statistics for retry results
                    retry_success_count = 0
                    for item_id, filename, hash_value, is_update, display_path in failed_items:
                        if retry_results.get(item_id, False):
                            retry_success_count += 1
                            # Correct the statistics: remove 1 from failed, add to succeeded
                            self.stats_wrapper.decrement('hash_save_failed')

                            if is_update:
                                self.stats_wrapper.increment('hash_updated')
                            else:
                                self.stats_wrapper.increment('hash_new_saved')

                    if retry_success_count > 0:
                        print(f"[✓] Retry successful for {retry_success_count}/{len(failed_items)} files")

                    # If some still failed, try one more time with even longer delay
                    if retry_success_count < len(failed_items):
                        still_failed_items = []
                        for item_id, filename, hash_value, is_update, display_path in failed_items:
                            if not retry_results.get(item_id, False):
                                still_failed_items.append((item_id, filename, hash_value, is_update, display_path))

                        if still_failed_items:
                            # Check what types of files are still failing
                            still_html = sum(1 for _, f, _, _, _ in still_failed_items if f.lower().endswith('.html'))
                            still_office_pdf = sum(1 for _, f, _, _, _ in still_failed_items
                                              if any(f.lower().endswith(ext) for ext in ['.pdf', '.docx', '.xlsx', '.pptx', '.doc', '.xls', '.ppt']))
                            still_images = sum(1 for _, f, _, _, _ in still_failed_items
                                             if any(f.lower().endswith(ext) for ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp', '.tiff']))
                            still_other = len(still_failed_items) - still_html - still_office_pdf - still_images

                            print(f"[⏱] {len(still_failed_items)} files still failing. Final retry in 20 seconds...")
                            if is_debug_enabled():
                                type_breakdown = []
                                if still_html > 0:
                                    type_breakdown.append(f"{still_html} HTML")
                                if still_office_pdf > 0:
                                    type_breakdown.append(f"{still_office_pdf} Office/PDF")
                                if still_images > 0:
                                    type_breakdown.append(f"{still_images} images")
                                if still_other > 0:
                                    type_breakdown.append(f"{still_other} other")
                                if type_breakdown:
                                    print(f"    ({', '.join(type_breakdown)})")
                            time.sleep(20)

                            print(f"[#] Final retry for {len(still_failed_items)} files...")
                            final_retry_batch = [(item_id, filename, hash_value, display_path)
                                                for item_id, filename, hash_value, _, display_path in still_failed_items]

                            try:
                                final_results = batch_update_filehash_fields(
                                    config.tenant_url, library_name, final_retry_batch,
                                    config.tenant_id, config.client_id, config.client_secret,
                                    config.login_endpoint, config.graph_endpoint, batch_size=5  # Even smaller batches
                                )

                                final_success_count = 0
                                for item_id, filename, hash_value, is_update, display_path in still_failed_items:
                                    if final_results.get(item_id, False):
                                        final_success_count += 1
                                        # Correct the statistics
                                        self.stats_wrapper.decrement('hash_save_failed')
                                        if is_update:
                                            self.stats_wrapper.increment('hash_updated')
                                        else:
                                            self.stats_wrapper.increment('hash_new_saved')

                                if final_success_count > 0:
                                    print(f"[✓] Final retry successful for {final_success_count}/{len(still_failed_items)} files")

                                final_failed = len(still_failed_items) - final_success_count
                                if final_failed > 0:
                                    print(f"[!] {final_failed} files still failed after all retries (SharePoint may need more time)")

                            except Exception as final_error:
                                print(f"[!] Final retry failed: {str(final_error)[:200]}")

                except Exception as retry_error:
                    print(f"[!] Retry batch update failed: {str(retry_error)[:200]}")

        except Exception as e:
            print(f"[!] Batch metadata update failed: {e}")
            # Mark all as failed
            for _ in batch:
                self.stats_wrapper.increment('hash_save_failed')
