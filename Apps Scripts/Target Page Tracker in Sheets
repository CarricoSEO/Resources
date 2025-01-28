// Target Page Tracker in Google Sheets
// Your Google Sheet should have the following columns in this order: Client Name, URL,	Title (Current), Title (Previous),	Meta Description (Current),	Meta Description (Previous),	H1 (Current), H1 (Previous)	Status Code (Current), Status Code (Previous)
// The names are not required - however it should be noted that the sheet will extract the URL from column B (2), and will input/move the data based on the above columns.

function updatePageData() {
  var sheetName = "Sheet1"; // Target sheet name - should be edited based on your the name of your sheet/tab.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    Logger.log("Error: Sheet '" + sheetName + "' not found.");
    return;
  }

  var data = sheet.getDataRange().getValues();
  var changes = [];

  for (var i = 1; i < data.length; i++) {
    var url = data[i][1]; // URL is now in Column B (Index 1 instead of 0)
    if (!url) continue; // Skip empty rows

    // Fetch On-Page SEO Data (Title, Meta, H1)
    var seoData = fetchSEODataWithRegex(url);
    
    // Fetch Status Code using the robust function
    var statusCode = getStatusCode(url);

    if (!seoData) continue;

    // Move previous data to history columns
    data[i][3] = data[i][2]; // Previous Page Title
    data[i][5] = data[i][4]; // Previous Meta Description
    data[i][7] = data[i][6]; // Previous H1
    data[i][9] = data[i][8]; // Previous Status Code

    // Update current data (with decoded HTML entities)
    data[i][2] = decodeHtmlEntities(seoData.title);
    data[i][4] = decodeHtmlEntities(seoData.metaDescription);
    data[i][6] = decodeHtmlEntities(seoData.h1);
    data[i][8] = statusCode;

    // Check for changes
    if (data[i][2] !== data[i][3]) changes.push(`Page Title changed for ${url}`);
    if (data[i][4] !== data[i][5]) changes.push(`Meta Description changed for ${url}`);
    if (data[i][6] !== data[i][7]) changes.push(`H1 Tag changed for ${url}`);
    if (data[i][8] !== data[i][9]) changes.push(`Status Code changed for ${url} (${data[i][9]} → ${data[i][8]})`);
  }

  sheet.getDataRange().setValues(data);

  if (changes.length > 0) {
    sendEmailReport(changes);
  }
}

function fetchSEODataWithRegex(url) {
  try {
    var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    var html = response.getContentText();

    // Extract title
    var titleMatch = html.match(/<title[^>]*>(.*?)<\/title>/i);
    var title = titleMatch ? titleMatch[1] : "N/A";

    // Extract meta description (allowing multi-line content)
    var metaMatch = html.match(/<meta\s+name=["']description["']\s+content=["']([\s\S]*?)["']/i);
    var metaDescription = metaMatch ? metaMatch[1].trim() : "N/A"; // Trim any extra spaces or newlines

    // Extract clean H1
    var h1 = extractCleanH1(html);

    return { title, metaDescription, h1 };
  } catch (e) {
    Logger.log("Error fetching SEO data for " + url + ": " + e.toString());
    return null;
  }
}

// Function to extract clean H1 (removing HTML comments and nested elements)
function extractCleanH1(html) {
  var h1Match = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
  if (!h1Match) return "N/A";

  var h1Content = h1Match[1];

  // Remove HTML comments
  h1Content = h1Content.replace(/<!--[\s\S]*?-->/g, '');

  // Remove all nested tags, keeping only plain text
  h1Content = h1Content.replace(/<[^>]+>/g, '');

  // Trim whitespace
  return h1Content.trim() || "N/A";
}

// Function to decode HTML entities like &amp; → &, &#039; → '
function decodeHtmlEntities(str) {
  if (!str) return "N/A";

  var entityMap = {
    '&amp;': '&',
    '&lt;': '<',
    '&gt;': '>',
    '&quot;': '"',
    '&#039;': "'",
    '&#39;': "'",
    '&apos;': "'",
    '&nbsp;': ' ',
    '&copy;': '©',
    '&reg;': '®'
  };

  return str.replace(/&[#0-9a-zA-Z]+;/g, function (entity) {
    return entityMap[entity] || entity; // Replace if found, otherwise keep original
  });
}

function sendEmailReport(changes) {
  var recipient = "email1@email.com, email2@email.com";
  var subject = "**Changes to Target Pages Found**";
  var body = "This is an automated report. The following changes were detected:\n\n" + changes.join("\n");

  MailApp.sendEmail(recipient, subject, body);
}

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
