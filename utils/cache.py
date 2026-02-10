"""
Caching utilities for the application
Provides caching for frequently accessed data with expiration
"""

import time
import functools
from datetime import datetime, timedelta
from utils.logger import logger


class CacheEntry:
    """Represents a cached value with expiration time"""
    
    def __init__(self, value, ttl_seconds):
        """
        Initialize cache entry
        
        Args:
            value: The value to cache
            ttl_seconds: Time to live in seconds
        """
        self.value = value
        self.created_at = time.time()
        self.ttl_seconds = ttl_seconds
    
    def is_expired(self):
        """Check if cache entry has expired"""
        elapsed = time.time() - self.created_at
        return elapsed > self.ttl_seconds
    
    def get(self):
        """Get value if not expired, None if expired"""
        if self.is_expired():
            return None
        return self.value


class SimpleCache:
    """Simple in-memory cache with expiration"""
    
    def __init__(self):
        """Initialize cache"""
        self._cache = {}
        self._hits = 0
        self._misses = 0
    
    def get(self, key):
        """
        Get value from cache
        
        Args:
            key: Cache key
            
        Returns:
            Cached value or None if expired/not found
        """
        if key not in self._cache:
            self._misses += 1
            return None
        
        entry = self._cache[key]
        if entry.is_expired():
            del self._cache[key]
            self._misses += 1
            logger.debug(f"Cache expired for key: {key}")
            return None
        
        self._hits += 1
        logger.debug(f"Cache hit for key: {key}")
        return entry.value
    
    def set(self, key, value, ttl_seconds=3600):
        """
        Set value in cache
        
        Args:
            key: Cache key
            value: Value to cache
            ttl_seconds: Time to live in seconds (default: 1 hour)
        """
        self._cache[key] = CacheEntry(value, ttl_seconds)
        logger.debug(f"Cache set for key: {key} (TTL: {ttl_seconds}s)")
    
    def clear(self):
        """Clear all cache"""
        self._cache.clear()
        logger.info("Cache cleared")
    
    def cleanup_expired(self):
        """Remove expired entries"""
        expired_keys = [
            key for key, entry in self._cache.items() 
            if entry.is_expired()
        ]
        for key in expired_keys:
            del self._cache[key]
        
        if expired_keys:
            logger.debug(f"Cleaned up {len(expired_keys)} expired cache entries")
        
        return len(expired_keys)
    
    def get_stats(self):
        """Get cache statistics"""
        total = self._hits + self._misses
        hit_rate = (self._hits / total * 100) if total > 0 else 0
        return {
            'hits': self._hits,
            'misses': self._misses,
            'total_requests': total,
            'hit_rate': hit_rate,
            'cached_items': len(self._cache)
        }


# Global cache instances
_exchange_rate_cache = SimpleCache()
_hotel_cache = SimpleCache()
_client_cache = SimpleCache()


def cached_exchange_rates(ttl_seconds=3600):
    """
    Decorator to cache exchange rates
    
    Args:
        ttl_seconds: Time to live in seconds (default: 1 hour)
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            cache_key = 'exchange_rates'
            
            # Check cache
            cached_value = _exchange_rate_cache.get(cache_key)
            if cached_value is not None:
                return cached_value
            
            # Get fresh data
            result = func(*args, **kwargs)
            
            # Cache it
            _exchange_rate_cache.set(cache_key, result, ttl_seconds)
            
            return result
        
        return wrapper
    
    return decorator


def _make_cache_key(prefix, args, kwargs):
    if not args and not kwargs:
        return prefix
    return (prefix, args, tuple(sorted(kwargs.items())))


def cached_hotel_data(ttl_seconds=86400):
    """
    Decorator to cache hotel data
    
    Args:
        ttl_seconds: Time to live in seconds (default: 24 hours)
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            cache_key = _make_cache_key('all_hotels', args, kwargs)
            
            # Check cache
            cached_value = _hotel_cache.get(cache_key)
            if cached_value is not None:
                return cached_value
            
            # Get fresh data
            result = func(*args, **kwargs)
            
            # Cache it
            _hotel_cache.set(cache_key, result, ttl_seconds)
            
            return result
        
        return wrapper
    
    return decorator


def cached_client_data(ttl_seconds=86400):
    """
    Decorator to cache client data
    
    Args:
        ttl_seconds: Time to live in seconds (default: 24 hours)
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            cache_key = _make_cache_key('all_clients', args, kwargs)
            
            # Check cache
            cached_value = _client_cache.get(cache_key)
            if cached_value is not None:
                return cached_value
            
            # Get fresh data
            result = func(*args, **kwargs)
            
            # Cache it
            _client_cache.set(cache_key, result, ttl_seconds)
            
            return result
        
        return wrapper
    
    return decorator


def invalidate_all_caches():
    """Invalidate all caches"""
    _exchange_rate_cache.clear()
    _hotel_cache.clear()
    _client_cache.clear()
    logger.info("All caches invalidated")


def invalidate_hotel_cache():
    """Invalidate hotel cache"""
    _hotel_cache.clear()
    logger.info("Hotel cache invalidated")


def invalidate_client_cache():
    """Invalidate client cache"""
    _client_cache.clear()
    logger.info("Client cache invalidated")


def get_cache_stats():
    """Get statistics for all caches"""
    return {
        'exchange_rates': _exchange_rate_cache.get_stats(),
        'hotels': _hotel_cache.get_stats(),
        'clients': _client_cache.get_stats()
    }
