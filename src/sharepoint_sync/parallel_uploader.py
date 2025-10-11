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
from .markdown_converter import convert_markdown_to_html
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

    def _upload_files_parallel(self, file_list, site_id, drive_id, root_item_id, base_path, config,
                               filehash_available, library_name):
        """
        Upload regular files in parallel.

        Returns:
            int: Number of failed uploads
        """
        failed_count = 0

        def upload_worker(filepath):
            """Worker function for parallel upload"""
            # Enable thread-safe print for this thread
            enable_thread_safe_print()

            try:
                # Call existing upload function - maintains all output/statistics
                upload_file_with_structure(
                    site_id, drive_id, root_item_id, filepath, base_path,
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

        # Execute uploads in parallel
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Submit all upload tasks
            future_to_file = {
                executor.submit(upload_worker, f): f
                for f in file_list
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

        def process_md_worker(md_filepath):
            """Worker for markdown processing"""
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
                executor.submit(process_md_worker, f): f
                for f in md_files
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
            # Read markdown content
            with open(file_path, 'r', encoding='utf-8') as md_file_handle:
                md_content = md_file_handle.read()

            # Convert to HTML
            html_content = convert_markdown_to_html(md_content, os.path.basename(file_path))

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

            # Upload HTML file
            for i in range(config.max_retry):
                try:
                    upload_file(
                        site_id, drive_id, target_folder_id, html_path, 4*1024*1024, config.force_upload,
                        config.tenant_url, library_name, filehash_available,
                        config.tenant_id, config.client_id, config.client_secret,
                        config.login_endpoint, config.graph_endpoint,
                        self.stats_wrapper, desired_name=desired_html_filename,
                        metadata_queue=self.metadata_queue  # Pass queue for batch updates
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
            self._process_metadata_batch(remaining, config, library_name)

    def _process_metadata_batch(self, batch, config, library_name):
        """
        Process batch of metadata updates.

        Args:
            batch (list): List of (item_id, filename, hash_value, is_file_update) tuples
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
        update_types = {item_id: is_update for item_id, _, _, is_update in batch}

        # Convert batch to format expected by batch_update_filehash_fields
        # It expects (item_id, filename, hash_value) tuples
        api_batch = [(item_id, filename, hash_value) for item_id, filename, hash_value, _ in batch]

        try:
            results = batch_update_filehash_fields(
                config.tenant_url, library_name, api_batch,
                config.tenant_id, config.client_id, config.client_secret,
                config.login_endpoint, config.graph_endpoint
            )

            # Update statistics based on results and update type
            for item_id, success in results.items():
                if success:
                    # Track based on whether this was new file or update
                    is_update = update_types.get(item_id, False)
                    if is_update:
                        self.stats_wrapper.increment('hash_updated')
                    else:
                        self.stats_wrapper.increment('hash_new_saved')
                else:
                    self.stats_wrapper.increment('hash_save_failed')

        except Exception as e:
            print(f"[!] Batch metadata update failed: {e}")
            # Mark all as failed
            for _ in batch:
                self.stats_wrapper.increment('hash_save_failed')
