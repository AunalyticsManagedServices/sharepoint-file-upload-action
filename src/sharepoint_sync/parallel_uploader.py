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

        # Thread-safe list for files with Mermaid diagram failures
        # Each item: (relative_path, num_failed, num_total)
        self.mermaid_failed_files = []
        self.mermaid_failed_files_lock = __import__('threading').Lock()

        # Queue for batch metadata updates
        self.metadata_queue = BatchQueue(batch_size=20) if self.batch_metadata else None

    def process_files(self, local_files, site_id, drive_id, root_item_id, base_path, config,
                     filehash_available, library_name, converted_md_files_set=None, sharepoint_cache=None):
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
            sharepoint_cache (dict): Optional pre-built cache of SharePoint file metadata

        Returns:
            int: Number of failed uploads
        """
        # Store cache for workers to access
        # Extract files cache from new structure if present
        if isinstance(sharepoint_cache, dict) and 'files' in sharepoint_cache:
            # New structure: {'files': {...}, 'folders': {...}}
            self.sharepoint_cache = sharepoint_cache['files']
            self.folder_cache = sharepoint_cache.get('folders')
        else:
            # Old structure: direct dict of files
            self.sharepoint_cache = sharepoint_cache
            self.folder_cache = None

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
            md_start_time = time.time()
            print(f"[*] Processing markdown files:")
            if is_debug_enabled():
                print(f"[DEBUG] Converting {len(md_files)} markdown files in parallel...")

            failed_count += self._process_markdown_files_parallel(
                md_files, site_id, drive_id, root_item_id, base_path, config,
                filehash_available, library_name
            )

            # Show detailed summary after markdown processing
            md_no_changes = self.stats_wrapper.get('md_no_changes', 0)
            md_converted = self.stats_wrapper.get('md_converted', 0)
            md_failed = self.stats_wrapper.get('md_conversion_failed', 0)
            mermaid_rendered = self.stats_wrapper.get('mermaid_diagrams_rendered', 0)
            mermaid_failed = self.stats_wrapper.get('mermaid_diagrams_failed', 0)
            total_md = len(md_files)

            print(f"   - No Changes Detected:        {md_no_changes:>4}/{total_md}")
            print(f"   - Converted to HTML:          {md_converted:>4}/{total_md}")
            if md_failed > 0:
                print(f"   - Conversion Failed:          {md_failed:>4}")

            # Show Mermaid diagram statistics if any diagrams were processed
            if mermaid_rendered > 0 or mermaid_failed > 0:
                total_mermaid = mermaid_rendered + mermaid_failed
                print(f"   - Mermaid Diagrams Rendered:  {mermaid_rendered:>4}/{total_mermaid}")
                if mermaid_failed > 0:
                    print(f"   - Mermaid Diagrams Failed:    {mermaid_failed:>4}/{total_mermaid}")

                    # Display list of files with Mermaid failures
                    if self.mermaid_failed_files:
                        print(f"\n   Files with Mermaid diagram failures:")
                        for file_path, num_failed, num_total in sorted(self.mermaid_failed_files):
                            print(f"      - {file_path} ({num_failed}/{num_total} diagrams failed)")

            md_elapsed = time.time() - md_start_time
            converted_count = len([f for f in md_files if f in self.converted_md_files])
            if converted_count > 0:
                print(f"\n[✓] Verified or converted {converted_count} markdown files ({md_elapsed:.3f}s)")

        # Process regular files in parallel
        upload_start_time = time.time()
        if regular_files:
            # Check if any files actually need uploading (not just skipped)
            files_before = self.stats_wrapper.get('new_files', 0) + self.stats_wrapper.get('replaced_files', 0)

            print(f"\n[*] Uploading files...")
            if is_debug_enabled():
                print(f"[DEBUG] Uploading {len(regular_files)} files in parallel (workers: {self.max_workers})...")

            failed_count += self._upload_files_parallel(
                regular_files, site_id, drive_id, root_item_id, base_path, config,
                filehash_available, library_name
            )

            upload_elapsed = time.time() - upload_start_time
            files_after = self.stats_wrapper.get('new_files', 0) + self.stats_wrapper.get('replaced_files', 0)
            files_uploaded = files_after - files_before

            if files_uploaded == 0:
                print(f"\n[✓] No file changes detected ({upload_elapsed:.3f}s)")
            else:
                print(f"\n[✓] Uploaded {files_uploaded} files ({upload_elapsed:.3f}s)")
        else:
            upload_elapsed = time.time() - upload_start_time
            print(f"\n[✓] No files to upload ({upload_elapsed:.3f}s)")

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
                    metadata_queue=self.metadata_queue,  # Pass queue for batch updates
                    sharepoint_cache=self.sharepoint_cache  # Pass cache for instant lookups
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
            from .file_handler import calculate_file_hash, check_file_needs_update
            md_file_hash = calculate_file_hash(file_path)
            if md_file_hash and is_debug_enabled():
                print(f"[#] Source .md file hash: {md_file_hash[:8]}... (will be used for .html file)")

            # Determine target .html filename and folder BEFORE converting
            # We need this to check if the file already exists with matching hash
            original_html_path = file_path.replace('.md', '.html')
            desired_html_filename = os.path.basename(original_html_path)

            # Calculate relative path and target folder
            if base_path:
                rel_path_str = os.path.relpath(original_html_path, base_path)
            else:
                rel_path_str = original_html_path

            # Normalize and sanitize
            if isinstance(rel_path_str, bytes):
                rel_path_str = rel_path_str.decode('utf-8')
            rel_path_str = rel_path_str.replace('\\', '/')
            sanitized_rel_path = sanitize_path_components(rel_path_str)
            dir_path = os.path.dirname(sanitized_rel_path)

            # Determine target folder ID
            target_folder_id = root_item_id
            if dir_path and dir_path != "." and dir_path != "":
                from .uploader import ensure_folder_exists

                # Use stored folder cache (already extracted in process_files)
                target_folder_id = ensure_folder_exists(
                    site_id, drive_id, root_item_id, dir_path,
                    config.tenant_id, config.client_id, config.client_secret,
                    config.login_endpoint, config.graph_endpoint,
                    folder_cache=self.folder_cache
                )

            # EARLY CHECK: Does .html file already exist with matching source .md hash?
            # This avoids expensive markdown conversion if source hasn't changed
            # Skip this check if force_md_to_html_regeneration is enabled
            if not config.force_upload and not config.force_md_to_html_regeneration and filehash_available:
                if is_debug_enabled():
                    print(f"[?] Checking if converted .html file needs update: {sanitized_rel_path}")

                needs_update, exists, _, _ = check_file_needs_update(
                    file_path,  # We pass .md file path for hash calculation (but hash already calculated)
                    desired_html_filename,  # Check for .html file in SharePoint
                    config.tenant_url, library_name, filehash_available,
                    config.tenant_id, config.client_id, config.client_secret,
                    config.login_endpoint, config.graph_endpoint,
                    self.stats_wrapper,
                    pre_calculated_hash=md_file_hash,  # Use source .md hash for comparison
                    display_path=sanitized_rel_path,
                    site_id=site_id, drive_id=drive_id, parent_item_id=target_folder_id,
                    sharepoint_cache=self.sharepoint_cache  # Use cache for instant lookup
                )

                if not needs_update:
                    # File exists and source .md hash matches - SKIP conversion entirely!
                    if is_debug_enabled():
                        print(f"[=] Skipping markdown conversion - source unchanged: {sanitized_rel_path}")
                    self.stats_wrapper.increment('md_no_changes')
                    return True  # Success - no work needed

            # File needs update or doesn't exist - proceed with conversion
            if is_debug_enabled():
                print(f"[MD] Converting markdown to HTML: {file_path}")

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
            html_content, mermaid_success, mermaid_failed = convert_markdown_to_html(
                md_content,
                file_path,
                sharepoint_base_url=sharepoint_base_url,
                current_file_rel_path=rel_path_str
            )

            # Track Mermaid diagram statistics
            if mermaid_success > 0:
                for _ in range(mermaid_success):
                    self.stats_wrapper.increment('mermaid_diagrams_rendered')
            if mermaid_failed > 0:
                for _ in range(mermaid_failed):
                    self.stats_wrapper.increment('mermaid_diagrams_failed')

                # Track which file had failures for detailed reporting
                total_mermaid = mermaid_success + mermaid_failed
                with self.mermaid_failed_files_lock:
                    self.mermaid_failed_files.append((sanitized_rel_path, mermaid_failed, total_mermaid))

            # Create temp HTML file
            temp_html_fd, html_path = tempfile.mkstemp(suffix='.html', prefix='converted_md_')

            try:
                with os.fdopen(temp_html_fd, 'w', encoding='utf-8') as html_file:
                    html_file.write(html_content)
            except Exception as write_error:
                os.close(temp_html_fd)
                raise write_error

            # Paths and target folder already calculated above (before early check)
            # No need to recalculate: original_html_path, desired_html_filename, sanitized_rel_path, target_folder_id

            # Upload HTML file with source .md file hash
            # This allows hash-based comparison instead of size-only (solves Mermaid SVG ID variation issue)
            # Force upload if force_md_to_html_regeneration is true (always upload newly regenerated HTML)
            force_html_upload = config.force_upload or config.force_md_to_html_regeneration
            for i in range(config.max_retry):
                try:
                    upload_file(
                        site_id, drive_id, target_folder_id, html_path, 4*1024*1024, force_html_upload,
                        config.tenant_url, library_name, filehash_available,
                        config.tenant_id, config.client_id, config.client_secret,
                        config.login_endpoint, config.graph_endpoint,
                        self.stats_wrapper, desired_name=desired_html_filename,
                        metadata_queue=self.metadata_queue,  # Pass queue for batch updates
                        pre_calculated_hash=md_file_hash,  # Use source .md file hash for comparison
                        display_path=sanitized_rel_path,  # Show full relative path in debug output
                        sharepoint_cache=self.sharepoint_cache  # Pass cache for instant lookups
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

            self.stats_wrapper.increment('md_converted')
            return True

        except Exception as e:
            print(f"[Error] Failed to convert markdown file {file_path}: {e}")
            self.stats_wrapper.increment('md_conversion_failed')
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
            if is_debug_enabled():
                print(f"[DEBUG] Metadata queue is empty - nothing to flush")
            return

        # Check queue size before processing
        queue_size = self.metadata_queue.qsize()
        if is_debug_enabled():
            print(f"[DEBUG] Metadata queue contains approximately {queue_size} items")

        print(f"[#] Processing remaining metadata updates...")

        # Get all remaining items
        remaining = self.metadata_queue.get_all_remaining()
        if is_debug_enabled():
            print(f"[DEBUG] Retrieved {len(remaining)} items from queue")

        if remaining:
            # Add delay for complex file types to allow SharePoint processing to complete
            # Different file types need processing time: virus scan, content indexing, conversion, sanitization
            html_count = sum(1 for _, filename, _, _, _, _ in remaining
                           if filename.lower().endswith('.html'))
            pdf_count = sum(1 for _, filename, _, _, _, _ in remaining
                          if filename.lower().endswith('.pdf'))
            office_count = sum(1 for _, filename, _, _, _, _ in remaining
                              if any(filename.lower().endswith(ext) for ext in ['.docx', '.xlsx', '.pptx', '.doc', '.xls', '.ppt']))
            image_count = sum(1 for _, filename, _, _, _, _ in remaining
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
            batch (list): List of (parent_item_id, filename, item_id, hash_value, is_update, display_path) tuples
            config: Configuration object
            library_name (str): SharePoint library name
        """
        # Import here to avoid circular dependency
        from .graph_api import batch_update_filehash_fields

        if not batch:
            return

        print(f"[#] Batch updating {len(batch)} FileHash values...")

        # Extract update type info for statistics tracking
        update_types = {}
        for parent_id, filename, item_id, _, is_update, _ in batch:
            update_types[item_id] = is_update

        # Convert batch to format expected by batch_update_filehash_fields
        # Format: (item_id, filename, hash_value, display_path)
        api_batch = [(item_id, filename, hash_value, display_path)
                     for _, filename, item_id, hash_value, _, display_path in batch]

        try:
            results = batch_update_filehash_fields(
                config.tenant_url, library_name, api_batch,
                config.tenant_id, config.client_id, config.client_secret,
                config.login_endpoint, config.graph_endpoint
            )

            # Update statistics based on results
            success_count = 0
            failed_items = []

            for parent_id, filename, item_id, hash_value, is_update, display_path in batch:
                success = results.get(item_id, False)

                if success:
                    success_count += 1
                    if is_update:
                        self.stats_wrapper.increment('hash_updated')
                    else:
                        self.stats_wrapper.increment('hash_new_saved')
                else:
                    # Store parent_id and filename for re-querying item_id on retry
                    failed_items.append((parent_id, filename, hash_value, is_update, display_path))
                    self.stats_wrapper.increment('hash_save_failed')

            if is_debug_enabled():
                print(f"[✓] Batch update: {success_count}/{len(batch)} succeeded")

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

                print(f"[⏱] {len(failed_items)} files need retry. Waiting {retry_delay} seconds...")
                time.sleep(retry_delay)

                print(f"[#] Retrying {len(failed_items)} failed FileHash updates (re-querying item IDs)...")

                # Re-query fresh item IDs for failed files only
                from .graph_api import get_drive_item_by_path_with_list_item
                retry_batch = []

                for parent_id, filename, hash_value, is_update, display_path in failed_items:
                    try:
                        # Query fresh item ID using path
                        # Note: We need site_id and drive_id - get from first successful query or config
                        # For now, rely on batch_update_filehash_fields to handle this
                        # Store as None to indicate it needs re-query
                        retry_batch.append((parent_id, filename, None, hash_value, is_update, display_path))
                    except Exception as e:
                        if is_debug_enabled():
                            print(f"[DEBUG] Failed to prepare retry for {display_path}: {str(e)[:100]}")

                try:
                    # Pass batch with None item_ids to indicate they need re-querying
                    retry_results = batch_update_filehash_fields(
                        config.tenant_url, library_name, retry_batch,
                        config.tenant_id, config.client_id, config.client_secret,
                        config.login_endpoint, config.graph_endpoint, batch_size=10,
                        requery_item_ids=True  # Signal to re-query item IDs
                    )

                    # Update statistics for retry results
                    retry_success_count = 0
                    for idx, (parent_id, filename, _, hash_value, is_update, display_path) in enumerate(retry_batch):
                        # Results keyed by original index or identifier
                        if retry_results.get(idx, False):
                            retry_success_count += 1
                            self.stats_wrapper.decrement('hash_save_failed')
                            if is_update:
                                self.stats_wrapper.increment('hash_updated')
                            else:
                                self.stats_wrapper.increment('hash_new_saved')

                    if retry_success_count > 0:
                        print(f"[✓] Retry successful for {retry_success_count}/{len(failed_items)} files")

                    # If some still failed, try one more time
                    if retry_success_count < len(failed_items):
                        still_failed = [
                            retry_batch[idx] for idx in range(len(retry_batch))
                            if not retry_results.get(idx, False)
                        ]

                        if still_failed:
                            print(f"[⏱] {len(still_failed)} files still failing. Final retry in 20 seconds...")
                            time.sleep(20)

                            print(f"[#] Final retry for {len(still_failed)} files...")

                            try:
                                final_results = batch_update_filehash_fields(
                                    config.tenant_url, library_name, still_failed,
                                    config.tenant_id, config.client_id, config.client_secret,
                                    config.login_endpoint, config.graph_endpoint, batch_size=5,
                                    requery_item_ids=True
                                )

                                final_success_count = sum(1 for success in final_results.values() if success)

                                if final_success_count > 0:
                                    print(f"[✓] Final retry successful for {final_success_count}/{len(still_failed)} files")
                                    # Correct statistics
                                    for _ in range(final_success_count):
                                        self.stats_wrapper.decrement('hash_save_failed')

                                final_failed = len(still_failed) - final_success_count
                                if final_failed > 0:
                                    print(f"[!] {final_failed} files still failed after all retries")

                            except Exception as final_error:
                                print(f"[!] Final retry failed: {str(final_error)[:200]}")

                except Exception as retry_error:
                    print(f"[!] Retry batch update failed: {str(retry_error)[:200]}")

        except Exception as e:
            print(f"[!] Batch metadata update failed: {e}")
            # Mark all as failed
            for _ in batch:
                self.stats_wrapper.increment('hash_save_failed')
