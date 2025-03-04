function updatePageData() {
  var sheetName = "AutomatedReportTest";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Error: Sheet '" + sheetName + "' not found.");
    return;
  }

  var data = sheet.getDataRange().getValues();
  var changesByClient = {};

  for (var i = 1; i < data.length; i++) {
    var client = data[i][0]; // Client Name (Column A)
    var url = data[i][1]; // URL (Column B)
    if (!url) continue;

    var seoData = fetchSEODataWithRegex(url);
    var statusCode = getStatusCode(url);
    if (!seoData) continue;

    data[i][3] = data[i][2]; // Previous Page Title
    data[i][5] = data[i][4]; // Previous Meta Description
    data[i][7] = data[i][6]; // Previous H1
    data[i][9] = data[i][8]; // Previous Status Code

    data[i][2] = decodeHtmlEntities(seoData.title);
    data[i][4] = decodeHtmlEntities(seoData.metaDescription);
    data[i][6] = decodeHtmlEntities(seoData.h1);
    data[i][8] = statusCode;

    var clientChanges = [];

    if (data[i][2] !== data[i][3]) {
      clientChanges.push(`<b>Title changed for</b> <a href='${url}'>${url}</a><br><b style='color:#0C232A;'>Old:</b> <s>"${data[i][3]}"</s> → <b style='color:#C9A82D;'>New:</b> "${data[i][2]}"`);
    }
    if (data[i][4] !== data[i][5]) {
      clientChanges.push(`<b>Meta Description changed for</b> <a href='${url}'>${url}</a><br><b style='color:#0C232A;'>Old:</b> <s>"${data[i][5]}"</s> → <b style='color:#C9A82D;'>New:</b> "${data[i][4]}"`);
    }
    if (data[i][6] !== data[i][7]) {
      clientChanges.push(`<b>H1 Tag changed for</b> <a href='${url}'>${url}</a><br><b style='color:#0C232A;'>Old:</b> <s>"${data[i][7]}"</s> → <b style='color:#C9A82D;'>New:</b> "${data[i][6]}"`);
    }
    if (String(data[i][8]) !== String(data[i][9])) {
      clientChanges.push(`<b>Status Code changed for</b> <a href='${url}'>${url}</a> (<b style='color:#0C232A;'>Old:</b> ${data[i][9]} → <b style='color:#C9A82D;'>New:</b> ${data[i][8]})`);
    }

    if (clientChanges.length > 0) {
      if (!changesByClient[client]) {
        changesByClient[client] = [];
      }
      changesByClient[client].push(...clientChanges);
    }
  }

  sheet.getDataRange().setValues(data);
  if (Object.keys(changesByClient).length > 0) {
    sendEmailReport(changesByClient);
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

function sendEmailReport(changesByClient) {
  var recipient = "example@email.com, example2@email.com"; // **** INSERT YOUR EMAIL(s) HERE
  var subject = "!!Changes to Target Pages Found!!";
  
  var htmlBody = `<p>The following page changes were detected:</p>`;
  
  for (var client in changesByClient) {
    htmlBody += `<h2>${client}</h2><ul>`;
    changesByClient[client].forEach(change => {
      htmlBody += `<li>${change}</li>`;
    });
    htmlBody += `</ul>`;
  }
// **** ADD YOUR GOOGLE SHEETS URL BELOW
  htmlBody += `<br>
  <i>This is an automated email report sent to you via the Google Apps Script trigger found in the following report:</i>
  <p><a href='INSERT_YOUR_GOOGLE_SHEETS_URL_HERE'
       style='background-color: #C9A82D; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;'>
       Target Page Tracker
    </a></p>
  <br><br>
  <div style="border: 2px solid #C9A82D; background-color: #0C232A; padding: 15px; border-radius: 10px; width: 30%; margin-left: 0; text-align: left;">
    <div style="width: 100%; display: flex; align-items: center;">
      <span style="color: white; font-size: 16px; margin-right: 10px;">
        Built by
      </span>
      <a href="https://www.carricoseo.com/?utm_source=google_apps_script&utm_medium=referral&utm_campaign=bannerad" target="_blank" style="text-decoration: none;">
        <img src="https://www.carricoseo.com/wp-content/uploads/2025/02/CarricoSEO-Logo-White.png" alt="CarricoSEO Logo" style="height: 25px; vertical-align: middle;">
      </a>
    </div>
    <br>
    <div style="margin-top: 10px; text-align: left; color: white; font-size: 14px;">
      This and many other one-click tools can be found at
      <a href="https://tools.carricoseo.com/?utm_source=google_apps_script&utm_medium=referral&utm_campaign=bannerad" target="_blank" style="color: lightblue; text-decoration: none;">
        CS Tools.
      </a>
      I also have a blog full of other free resources, tools, and scripts
      <a href="https://www.carricoseo.com/resources/?utm_source=google_colab&utm_medium=referral&utm_campaign=colab_blog" target="_blank" style="color: lightblue; text-decoration: none;">
        found here!
      </a>
    </div>
  </div>`;

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlBody
  });
}
