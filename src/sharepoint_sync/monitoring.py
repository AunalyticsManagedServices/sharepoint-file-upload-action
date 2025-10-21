# -*- coding: utf-8 -*-
"""
Rate limiting monitoring and statistics tracking for SharePoint sync.

This module provides classes for monitoring Graph API rate limits and tracking
upload statistics.
"""

import os


class RateLimitMonitor:
    """
    Monitor and track Graph API rate limiting metrics.

    Analyzes response headers to detect and track throttling:
    - x-ms-throttle-limit-percentage: Utilization percentage (0.8-1.8 range)
    - x-ms-resource-unit: Resource units consumed per request
    - x-ms-throttle-scope: Throttling scope details

    Headers only appear when >80% of limit consumed.
    """

    def __init__(self):
        """Initialize rate limit monitoring metrics"""
        self.metrics = {
            'total_requests': 0,
            'throttled_requests': 0,
            'average_throttle_percentage': 0.0,
            'max_throttle_percentage': 0.0,
            'resource_units_consumed': 0,
            'alerts_triggered': 0
        }
        self.throttle_threshold = 0.8  # Alert when >80% of limit

        # Track API request types
        self.request_types = {
            'GET': 0,
            'POST': 0,
            'PUT': 0,
            'PATCH': 0,
            'DELETE': 0
        }

        # Track API operation types
        self.operations = {
            'file_upload': 0,           # PUT to /content endpoint
            'file_delete': 0,           # DELETE file
            'metadata_get': 0,          # GET file/folder metadata
            'metadata_update': 0,       # PATCH list item fields
            'folder_create': 0,         # POST create folder
            'batch_operation': 0,       # POST to $batch endpoint
            'cache_build': 0,           # GET with $expand for caching
            'column_ops': 0,            # GET/POST to /columns endpoint
            'site_list_info': 0,        # GET to /sites/ or /lists/ (no /items/)
            'other': 0                  # Other unclassified operations
        }

    def analyze_response_headers(self, response, method=None, url=None):
        """
        Analyze Graph API response headers for rate limiting info.

        Args:
            response: requests.Response object from Graph API call
            method (str): HTTP method (GET, POST, PUT, PATCH, DELETE)
            url (str): Request URL for operation type detection

        Returns:
            dict: Rate limiting information extracted from headers
        """
        self.metrics['total_requests'] += 1

        # Track request method type
        if method and method.upper() in self.request_types:
            self.request_types[method.upper()] += 1

        # Track operation type based on URL and method
        if url and method:
            self._categorize_operation(url, method.upper())

        headers = response.headers
        throttle_percentage = headers.get('x-ms-throttle-limit-percentage')
        resource_unit = headers.get('x-ms-resource-unit')
        throttle_scope = headers.get('x-ms-throttle-scope')

        if throttle_percentage:
            percentage = float(throttle_percentage)
            self.metrics['max_throttle_percentage'] = max(
                self.metrics['max_throttle_percentage'],
                percentage
            )

            # Calculate running average
            current_avg = self.metrics['average_throttle_percentage']
            total_requests = self.metrics['total_requests']
            self.metrics['average_throttle_percentage'] = (
                ((current_avg * (total_requests - 1)) + percentage) / total_requests
            )

            if percentage >= 1.0:
                self.metrics['throttled_requests'] += 1
                print(f"[!] THROTTLING DETECTED: {percentage:.1%} of limit used")

                if throttle_scope:
                    print(f"[!] Throttle scope: {throttle_scope}")

            elif percentage >= self.throttle_threshold:
                self.metrics['alerts_triggered'] += 1
                print(f"[ ] Rate limit warning: {percentage:.1%} of limit used")

        if resource_unit:
            units = int(resource_unit)
            self.metrics['resource_units_consumed'] += units
            # Only print if debug mode is enabled
            debug_metadata = os.environ.get('DEBUG_METADATA', 'false').lower() == 'true'
            if debug_metadata:
                print(f"[=] Resource units consumed: {units}")

        return {
            'throttle_percentage': float(throttle_percentage) if throttle_percentage else None,
            'resource_unit': int(resource_unit) if resource_unit else None,
            'throttle_scope': throttle_scope,
            'is_throttled': response.status_code == 429
        }

    def _categorize_operation(self, url, method):
        """
        Categorize API operation based on URL pattern and HTTP method.

        Args:
            url (str): Request URL
            method (str): HTTP method (GET, POST, PUT, PATCH, DELETE)
        """
        url_lower = url.lower()

        # File upload operations
        if method == 'PUT' and '/content' in url_lower:
            self.operations['file_upload'] += 1
        # File delete operations
        elif method == 'DELETE' and '/items/' in url_lower:
            self.operations['file_delete'] += 1
        # Metadata update operations
        elif method == 'PATCH' and '/listitem' in url_lower:
            self.operations['metadata_update'] += 1
        # Folder creation
        elif method == 'POST' and '/children' in url_lower:
            self.operations['folder_create'] += 1
        # Batch operations
        elif method == 'POST' and '$batch' in url_lower:
            self.operations['batch_operation'] += 1
        # Cache building operations (GET with $expand)
        elif method == 'GET' and '$expand=listitem' in url_lower:
            self.operations['cache_build'] += 1
        # Column operations (checking/creating FileHash column)
        elif '/columns' in url_lower:
            self.operations['column_ops'] += 1
        # Site/List info queries (not items or drives)
        elif method == 'GET' and ('/sites/' in url_lower or '/lists/' in url_lower) and '/items/' not in url_lower and '/drives/' not in url_lower:
            self.operations['site_list_info'] += 1
        # Metadata retrieval operations
        elif method == 'GET' and ('/items/' in url_lower or '/drives/' in url_lower):
            self.operations['metadata_get'] += 1
        # Other operations
        else:
            self.operations['other'] += 1

    def get_metrics_summary(self):
        """
        Get comprehensive rate limiting metrics.

        Returns:
            dict: Summary of all rate limiting metrics
        """
        return {
            'total_requests': self.metrics['total_requests'],
            'throttled_requests': self.metrics['throttled_requests'],
            'throttle_rate': self.metrics['throttled_requests'] / max(self.metrics['total_requests'], 1),
            'average_throttle_percentage': self.metrics['average_throttle_percentage'],
            'max_throttle_percentage': self.metrics['max_throttle_percentage'],
            'resource_units_consumed': self.metrics['resource_units_consumed'],
            'alerts_triggered': self.metrics['alerts_triggered']
        }

    def should_slow_down(self):
        """
        Determine if requests should be slowed down proactively.

        Returns:
            bool: True if approaching rate limits (>90% utilization)
        """
        return self.metrics['max_throttle_percentage'] >= 0.9


# Global rate limit monitor instance
rate_monitor = RateLimitMonitor()


def print_rate_limiting_summary():
    """
    Print comprehensive rate limiting statistics collected during execution.

    Displays:
    - Total API requests made
    - Number of throttled requests
    - Average and maximum throttle percentages
    - Resource units consumed
    - Alerts triggered

    Color-coded status based on throttling severity.
    """
    metrics = rate_monitor.get_metrics_summary()

    print("\n" + "="*60)
    print("GRAPH API RATE LIMITING SUMMARY")
    print("="*60)
    print(f"[STATS] API Request Statistics:")
    print(f"   - Total API Requests:       {metrics['total_requests']:>6}")
    print(f"   - Throttled Requests:       {metrics['throttled_requests']:>6} ({metrics['throttle_rate']:.1%})")
    print(f"   - Average Throttle %:       {metrics['average_throttle_percentage']:>6.1%}")
    print(f"   - Max Throttle %:           {metrics['max_throttle_percentage']:>6.1%}")
    print(f"   - Resource Units Used:      {metrics['resource_units_consumed']:>6}")
    print(f"   - Alerts Triggered:         {metrics['alerts_triggered']:>6}")

    # Request method breakdown
    if any(rate_monitor.request_types.values()):
        print(f"\n[API] Request Methods:")
        for method, count in rate_monitor.request_types.items():
            if count > 0:
                print(f"   - {f'{method} requests:':<27} {count:>6}")

    # Operation type breakdown
    if any(rate_monitor.operations.values()):
        print(f"\n[OPS] Operation Types:")
        for op_type, count in rate_monitor.operations.items():
            if count > 0:
                # Format operation name nicely
                op_name = op_type.replace('_', ' ').title()
                print(f"   - {f'{op_name}:':<27} {count:>6}")

    # Status indicator based on throttling severity
    if metrics['max_throttle_percentage'] >= 1.0:
        print(f"\n[!] WARNING: Hit throttling limits during execution")
    elif metrics['max_throttle_percentage'] >= 0.8:
        print(f"\n[ ] CAUTION: Approached throttling limits")
    else:
        print(f"\n[OK] Stayed within throttling limits")
    print("="*60)


class UploadStatistics:
    """Track upload statistics for sync operations"""

    def __init__(self):
        """Initialize upload statistics"""
        self.stats = {
            'new_files': 0,
            'replaced_files': 0,
            'skipped_files': 0,
            'failed_files': 0,
            'deleted_files': 0,
            'bytes_uploaded': 0,
            'bytes_skipped': 0,
            # File comparison method statistics
            'compared_by_hash': 0,
            'compared_by_size': 0,
            # FileHash column operation statistics
            'hash_new_saved': 0,      # New files with hash saved
            'hash_updated': 0,         # Existing files with hash updated
            'hash_matched': 0,         # Files skipped due to hash match
            'hash_save_failed': 0,     # Failed to save hash to SharePoint
            'hash_empty_found': 0,     # Files with empty FileHash (column exists but value is None)
            'hash_column_unavailable': 0,  # Files checked when FileHash column doesn't exist
            'hash_backfilled': 0,      # Files with hash backfilled (not re-uploaded)
            'hash_backfill_failed': 0,  # Failed backfill attempts
            # Cache performance statistics
            'cache_hits': 0,          # Successful cache lookups (avoided API call)
            'cache_misses': 0,        # Files not in cache (new files)
            'api_queries': 0,         # API queries needed (fallback when cache unavailable)
            # Markdown conversion statistics
            'md_no_changes': 0,       # Markdown files checked but unchanged (skipped)
            'md_converted': 0,        # Markdown files actually converted to HTML
            'md_conversion_failed': 0,  # Markdown files that failed conversion
            # Mermaid diagram statistics
            'mermaid_diagrams_rendered': 0,  # Mermaid diagrams successfully converted to SVG
            'mermaid_diagrams_failed': 0     # Mermaid diagrams that failed (shown as code blocks)
        }

    def print_summary(self, total_files, whatif_mode=False):
        """
        Print final summary report of upload statistics.

        Args:
            total_files (int): Total number of files processed
            whatif_mode (bool): Whether sync deletion is in WhatIf mode
        """
        print(f"[STATS] Sync Statistics:")
        print(f"   - New files uploaded:       {self.stats['new_files']:>6}")
        print(f"   - Files updated:            {self.stats['replaced_files']:>6}")
        print(f"   - Files skipped (unchanged):{self.stats['skipped_files']:>6}")

        # Show deleted files with WhatIf indicator if applicable
        if self.stats['deleted_files'] > 0:
            if whatif_mode:
                print(f"   - Files deleted (WhatIf):   {self.stats['deleted_files']:>6}")
            else:
                print(f"   - Files deleted:            {self.stats['deleted_files']:>6}")

        print(f"   - Failed uploads:           {self.stats['failed_files']:>6}")
        print(f"   - Total files processed:    {total_files:>6}")

        # File comparison method statistics
        total_comparisons = self.stats['compared_by_hash'] + self.stats['compared_by_size']
        if total_comparisons > 0:
            print(f"\n[COMPARE] File Comparison Methods:")
            print(f"   - Compared by hash:         {self.stats['compared_by_hash']:>6} ({self.stats['compared_by_hash']/total_comparisons*100:.1f}%)")
            print(f"   - Compared by size:         {self.stats['compared_by_size']:>6} ({self.stats['compared_by_size']/total_comparisons*100:.1f}%)")

        # FileHash column operation statistics
        total_hash_ops = (self.stats['hash_new_saved'] + self.stats['hash_updated'] +
                         self.stats['hash_matched'] + self.stats['hash_save_failed'] +
                         self.stats['hash_empty_found'] + self.stats['hash_column_unavailable'] +
                         self.stats['hash_backfilled'] + self.stats['hash_backfill_failed'])
        if total_hash_ops > 0:
            print(f"\n[HASH] FileHash Column Statistics:")
            if self.stats['hash_new_saved'] > 0:
                print(f"   - New hashes saved:         {self.stats['hash_new_saved']:>6}")
            if self.stats['hash_updated'] > 0:
                print(f"   - Hashes updated:           {self.stats['hash_updated']:>6}")
            if self.stats['hash_matched'] > 0:
                print(f"   - Hash matches (skipped):   {self.stats['hash_matched']:>6}")
            if self.stats['hash_backfilled'] > 0:
                print(f"   - Hashes backfilled:        {self.stats['hash_backfilled']:>6}")
            if self.stats['hash_empty_found'] > 0:
                print(f"   - Empty hash found:         {self.stats['hash_empty_found']:>6}")
            if self.stats['hash_column_unavailable'] > 0:
                print(f"   - Column unavailable:       {self.stats['hash_column_unavailable']:>6}")
            if self.stats['hash_save_failed'] > 0:
                print(f"   - Hash save failures:       {self.stats['hash_save_failed']:>6}")
            if self.stats['hash_backfill_failed'] > 0:
                print(f"   - Backfill failures:        {self.stats['hash_backfill_failed']:>6}")

        # Show cache performance statistics if cache was used
        total_cache_ops = self.stats.get('cache_hits', 0) + self.stats.get('cache_misses', 0)
        if total_cache_ops > 0:
            print(f"\n[CACHE] Cache Performance:")
            cache_hits = self.stats.get('cache_hits', 0)
            cache_misses = self.stats.get('cache_misses', 0)
            api_queries = self.stats.get('api_queries', 0)

            print(f"   - Cache hits:               {cache_hits:>6}")
            print(f"   - Cache misses:             {cache_misses:>6}")
            if api_queries > 0:
                print(f"   - API queries (fallback):   {api_queries:>6}")

            # Calculate cache efficiency
            if total_cache_ops > 0:
                cache_efficiency = (cache_hits / total_cache_ops) * 100
                print(f"   - Cache efficiency:         {cache_efficiency:>5.1f}% (API calls avoided)")

        print(f"\n[DATA] Transfer Summary:")
        print(f"   - Data uploaded:   {format_bytes(self.stats['bytes_uploaded'])}")
        print(f"   - Data skipped:    {format_bytes(self.stats['bytes_skipped'])}")
        print(f"   - Total savings:   {format_bytes(self.stats['bytes_skipped'])} ({self.stats['skipped_files']} files not re-uploaded)")

        # Calculate efficiency percentage
        total_bytes = self.stats['bytes_uploaded'] + self.stats['bytes_skipped']
        if total_bytes > 0:
            efficiency = (self.stats['bytes_skipped'] / total_bytes) * 100
            print(f"   - Sync efficiency: {efficiency:.1f}% (bandwidth saved by smart sync)")


def format_bytes(bytes_value):
    """
    Convert bytes to human-readable format.

    Args:
        bytes_value (int): Number of bytes to format

    Returns:
        str: Human-readable string (e.g., "1.5 MB")
    """
    for unit in ['B', 'KB', 'MB', 'GB']:
        if bytes_value < 1024.0:
            return f"{bytes_value:.1f} {unit}"
        bytes_value /= 1024.0
    return f"{bytes_value:.1f} TB"


# Global upload statistics instance
upload_stats = UploadStatistics()
