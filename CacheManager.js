/**
 * Manages caching operations for the add-on.
 */
const CacheManager = {
  // Cache expiration time in seconds (e.g., 1 hour)
  CACHE_EXPIRATION: 3600,

  /**
   * Gets a value from the cache.
   * @param {string} key The cache key.
   * @return {*} The cached value, or null if not found.
   */
  get: function(key) {
    const cache = CacheService.getUserCache();
    const value = cache.get(key);
    return value ? JSON.parse(value) : null;
  },

  /**
   * Sets a value in the cache.
   * @param {string} key The cache key.
   * @param {*} value The value to cache.
   */
  set: function(key, value) {
    const cache = CacheService.getUserCache();
    cache.put(key, JSON.stringify(value), this.CACHE_EXPIRATION);
  },

  /**
   * Removes a value from the cache.
   * @param {string} key The cache key.
   */
  remove: function(key) {
    const cache = CacheService.getUserCache();
    cache.remove(key);
  },

  /**
   * Clears all cached data for the current user.
   */
  clear: function() {
    const cache = CacheService.getUserCache();
    cache.removeAll();
  }
};