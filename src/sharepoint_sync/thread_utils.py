# -*- coding: utf-8 -*-
"""
Thread-safe utilities for parallel processing.

This module provides thread-safe wrappers for console output and statistics
to maintain compatibility with existing code while enabling parallel execution.
"""

import threading
import builtins
from queue import Queue, Empty

# Global locks for thread-safe operations
_console_lock = threading.Lock()
_original_print = builtins.print


def thread_safe_print(*args, **kwargs):
    """
    Thread-safe replacement for print() that ensures sequential output.
    Maintains exact same output format as standard print().

    Args:
        *args: Same as print()
        **kwargs: Same as print()
    """
    with _console_lock:
        _original_print(*args, **kwargs)


def enable_thread_safe_print():
    """
    Replace built-in print() with thread-safe version in current thread.
    Call this at the start of each worker thread.
    """
    builtins.print = thread_safe_print


def restore_original_print():
    """Restore original print() function"""
    builtins.print = _original_print


class ThreadSafeStatsWrapper:
    """
    Thread-safe wrapper for upload_stats.stats dictionary.

    Provides dictionary-like interface with automatic locking,
    maintaining 100% compatibility with existing code that accesses
    upload_stats.stats directly.

    Example:
        from sharepoint_sync.monitoring import upload_stats
        stats_wrapper = ThreadSafeStatsWrapper(upload_stats.stats)
        stats_wrapper['new_files'] += 1  # Thread-safe
        value = stats_wrapper.get('skipped_files', 0)  # Thread-safe
    """

    def __init__(self, stats_dict):
        """
        Initialize a wrapper around existing stats dictionary.

        Args:
            stats_dict (dict): Reference to upload_stats.stats dictionary
        """
        self._stats = stats_dict  # Reference to actual stats dict
        self._lock = threading.Lock()

    def __getitem__(self, key):
        """Thread-safe dictionary access: stats[key]"""
        with self._lock:
            return self._stats[key]

    def __setitem__(self, key, value):
        """Thread-safe dictionary assignment: stats[key] = value"""
        with self._lock:
            self._stats[key] = value

    def get(self, key, default=None):
        """Thread-safe dictionary get: stats.get(key, default)"""
        with self._lock:
            return self._stats.get(key, default)

    def __contains__(self, key):
        """Thread-safe containment check: key in stats"""
        with self._lock:
            return key in self._stats

    def increment(self, key, value=1):
        """
        Thread-safe increment operation.

        Args:
            key (str): Statistics field to increment
            value (int/float): Amount to increment by (default: 1)
        """
        with self._lock:
            self._stats[key] = self._stats.get(key, 0) + value

    def add_bytes(self, key, bytes_count):
        """
        Thread-safe byte counter update.

        Args:
            key (str): Byte counter field ('bytes_uploaded' or 'bytes_skipped')
            bytes_count (int): Number of bytes to add
        """
        with self._lock:
            self._stats[key] = self._stats.get(key, 0) + bytes_count


class ThreadSafeCounter:
    """
    Thread-safe counter for tracking operations.

    Example:
        >>> counter = ThreadSafeCounter()
        >>> counter.increment()
        >>> count = counter.value()
    """

    def __init__(self, initial=0):
        """
        Initialize counter.

        Args:
            initial (int): Initial counter value
        """
        self._value = initial
        self._lock = threading.Lock()

    def increment(self, amount=1):
        """
        Increment counter by specified amount.

        Args:
            amount (int): Amount to increment (default: 1)

        Returns:
            int: New counter value
        """
        with self._lock:
            self._value += amount
            return self._value

    def decrement(self, amount=1):
        """
        Decrement counter by specified amount.

        Args:
            amount (int): Amount to decrement (default: 1)

        Returns:
            int: New counter value
        """
        with self._lock:
            self._value -= amount
            return self._value

    def value(self):
        """
        Get current counter value.

        Returns:
            int: Current value
        """
        with self._lock:
            return self._value

    def reset(self):
        """Reset counter to zero"""
        with self._lock:
            self._value = 0


class ThreadSafeSet:
    """
    Thread-safe set for tracking converted files, processed items, etc.

    Example:
        >>> file_set = ThreadSafeSet()
        >>> file_set.add('file.md')
        >>> if 'file.md' in file_set:
        ...     print("Already processed")
    """

    def __init__(self):
        """Initialize empty thread-safe set"""
        self._set = set()
        self._lock = threading.Lock()

    def add(self, item):
        """
        Add item to set.

        Args:
            item: Item to add
        """
        with self._lock:
            self._set.add(item)

    def remove(self, item):
        """
        Remove item from set.

        Args:
            item: Item to remove

        Raises:
            KeyError: If item not in set
        """
        with self._lock:
            self._set.remove(item)

    def discard(self, item):
        """
        Remove item from set if present.

        Args:
            item: Item to discard
        """
        with self._lock:
            self._set.discard(item)

    def __contains__(self, item):
        """Check if item in set (thread-safe)"""
        with self._lock:
            return item in self._set

    def __len__(self):
        """Get set size (thread-safe)"""
        with self._lock:
            return len(self._set)

    def copy(self):
        """
        Get copy of set contents.

        Returns:
            set: Copy of internal set
        """
        with self._lock:
            return self._set.copy()


class BatchQueue:
    """
    Thread-safe queue for collecting items to process in batches.

    This is particularly useful for batch metadata updates where multiple
    upload threads queue metadata changes, and a separate processor handles
    them in batches to reduce API calls.

    Example:
        >>> queue = BatchQueue(batch_size=20)
        >>> # Upload threads add items
        >>> queue.put(('item1', 'hash1'))
        >>> queue.put(('item2', 'hash2'))
        >>> # Processor thread gets batches
        >>> batch = queue.get_batch(timeout=5)
        >>> if batch:
        ...     # Process batch with your batch handler
        ...     for item in batch:
        ...         pass  # Handle item
    """

    def __init__(self, batch_size=20, max_wait_time=5.0):
        """
        Initialize batch queue.

        Args:
            batch_size (int): Maximum items per batch (default: 20 for Graph API)
            max_wait_time (float): Maximum seconds to wait for full batch
        """
        self._queue = Queue()
        self._batch_size = batch_size
        self._max_wait_time = max_wait_time
        self._closed = False
        self._lock = threading.Lock()

    def put(self, item):
        """
        Add item to queue.

        Args:
            item: Item to add (typically tuple of metadata to update)

        Raises:
            ValueError: If queue is closed
        """
        with self._lock:
            if self._closed:
                raise ValueError("Cannot put items in closed queue")
        self._queue.put(item)

    def get_batch(self, timeout=None):
        """
        Get batch of items from queue.

        Collects items until batch_size reached or timeout expires.

        Args:
            timeout (float): Maximum seconds to wait (default: max_wait_time)

        Returns:
            list: Batch of items (may be less than batch_size)
                  Empty list if timeout and no items available
        """
        if timeout is None:
            timeout = self._max_wait_time

        batch = []
        remaining_time = timeout

        import time
        start_time = time.time()

        while len(batch) < self._batch_size and remaining_time > 0:
            try:
                # Use shorter timeout for subsequent items
                item_timeout = min(remaining_time, 0.1) if batch else remaining_time
                item = self._queue.get(timeout=item_timeout)
                batch.append(item)

                # Update remaining time
                elapsed = time.time() - start_time
                remaining_time = timeout - elapsed

            except Empty:
                # No more items available within timeout
                break

        return batch

    def get_all_remaining(self):
        """
        Get all remaining items from queue without waiting.

        Useful for final cleanup when closing.

        Returns:
            list: All remaining items
        """
        items = []
        while not self._queue.empty():
            try:
                items.append(self._queue.get_nowait())
            except Empty:
                break
        return items

    def close(self):
        """Mark queue as closed (no more puts allowed)"""
        with self._lock:
            self._closed = True

    def is_closed(self):
        """Check if queue is closed"""
        with self._lock:
            return self._closed

    def qsize(self):
        """
        Get approximate queue size.

        Returns:
            int: Number of items in queue
        """
        return self._queue.qsize()

    def empty(self):
        """
        Check if queue is empty.

        Returns:
            bool: True if empty
        """
        return self._queue.empty()
