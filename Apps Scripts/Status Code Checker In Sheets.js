// Built by CarricoSEO
// You can find more information about this script @ 
// https://www.carricoseo.com/resources/status-code-check-in-google-sheets/

function getStatusCode(url) {
  // Trim any leading/trailing whitespace from the URL
  var url_trimmed = url.trim();

  // Handle empty URLs
  if (url_trimmed === '') {
    return 'Error: Empty URL';
  }

  // Get the script cache for storing results
  var cache = CacheService.getScriptCache();
  var cacheKey = 'status-' + url_trimmed;

  try {
    // Check if the cache key is too long (limit is 250 characters)
    if (cacheKey.length > 250) {
      return 'Error: URL too long for caching';
    }

    // Check if the result is already in the cache
    var cachedData = cache.get(cacheKey);

    if (cachedData) {
      return cachedData; // Return the cached status code or error message
    }

    try {
      // Configure options for UrlFetchApp
      var options = {
        'muteHttpExceptions': true, // Don't throw exceptions on HTTP errors
        'followRedirects': false, // Don't follow redirects
        'connectTimeout': 5000 // 5 seconds timeout
      };

      // Fetch the URL
      var response = UrlFetchApp.fetch(url_trimmed, options);
      var responseCode = response.getResponseCode();

      // Store the status code as a string
      var result = responseCode.toString();

      // Cache the result for 6 hours (21600 seconds)
      cache.put(cacheKey, result, 21600); // Cache the status code string

      // Return the status code
      return result;

    } catch (error) {
      // Enhanced Error Handling
      var errorMessage = 'Error: Unable to fetch URL';

      // Check for specific error types and provide more informative messages
      if (error.message.includes('DNS')) {
        errorMessage = 'Error: Hostname not found (DNS error)';
      } else if (error.message.includes('Timeout')) {
        errorMessage = 'Error: Connection timeout';
      } else if (error.message.includes('Unsupported')) {
        errorMessage = 'Error: Unsupported URL format';
      } else if (error.message.includes('Invalid argument')) {
        errorMessage = 'Error: Invalid URL';
      } else if (error.message.includes('Forbidden')) {
        errorMessage = 'Error: Access forbidden (403)';
      } else if (error.message.includes('Unauthorized')) {
        errorMessage = 'Error: Unauthorized access (401)';
      } else if (error.message.includes('Not Found')) {
        errorMessage = 'Error: Page not found (404)';
      } else if (error.message.includes('Internal')) {
        errorMessage = 'Error: Internal server error (500)';
      } else {
        // Include part of the original error message for debugging
        errorMessage += ' - ' + error.message.substring(0, 100);
      }

      return errorMessage; // Return the error message
    }

  } catch (error) {
    // Handle cache-related errors
    return 'Error: Cache issue - ' + error.message;
  }
}
