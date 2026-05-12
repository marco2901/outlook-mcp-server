/**
 * Microsoft Graph API helper functions
 */
const https = require('https');
const config = require('../config');
const mockData = require('./mock-data');

/**
 * Makes a request to the Microsoft Graph API
 * @param {string} accessToken - The access token for authentication
 * @param {string} method - HTTP method (GET, POST, etc.)
 * @param {string} path - API endpoint path
 * @param {object} data - Data to send for POST/PUT requests
 * @param {object} queryParams - Query parameters
 * @returns {Promise<object>} - The API response
 */
async function callGraphAPI(accessToken, method, path, data = null, queryParams = {}) {
  // For test tokens, we'll simulate the API call
  if (config.USE_TEST_MODE && accessToken.startsWith('test_access_token_')) {
    console.error(`TEST MODE: Simulating ${method} ${path} API call`);
    return mockData.simulateGraphAPIResponse(method, path, data, queryParams);
  }

  try {
    console.error(`Making real API call: ${method} ${path}`);
    
    // Check if path already contains the full URL (from nextLink)
    let finalUrl;
    if (path.startsWith('http://') || path.startsWith('https://')) {
      // Path is already a full URL (from pagination nextLink)
      finalUrl = path;
      console.error(`Using full URL from nextLink: ${finalUrl}`);
    } else {
      // Build URL from path and queryParams.
      // Path segments are percent-encoded ONLY if the caller didn't already
      // do it themselves. This avoids double-encoding IDs like base64-style
      // Outlook EwsIds (which contain `=`, `+`, `/`) that callers already
      // pass through encodeURIComponent — double-encoding produces literal
      // `%25` in the URL which Graph's todo endpoints reject with
      // RequestBroker--ParseUri.
      const encodedPath = path.split('/')
        .map((segment) =>
          /%[0-9A-Fa-f]{2}/.test(segment) ? segment : encodeURIComponent(segment)
        )
        .join('/');
      
      // Build query string from parameters. Keep OData parameter keys (e.g.
      // `$select`, `$top`, `$filter`) un-encoded — URLSearchParams would turn
      // `$` into `%24`, which most Graph endpoints accept but /me/todo/*
      // rejects with RequestBroker--ParseUri. Only the *value* is encoded.
      let queryString = '';
      if (Object.keys(queryParams).length > 0) {
        const filter = queryParams.$filter;
        if (filter) {
          delete queryParams.$filter;
        }

        const parts = [];
        for (const [key, value] of Object.entries(queryParams)) {
          parts.push(`${key}=${encodeURIComponent(value)}`);
        }
        queryString = parts.join('&');

        if (filter) {
          if (queryString) {
            queryString += `&$filter=${encodeURIComponent(filter)}`;
          } else {
            queryString = `$filter=${encodeURIComponent(filter)}`;
          }
        }

        if (queryString) {
          queryString = '?' + queryString;
        }

        console.error(`Query string: ${queryString}`);
      }
      
      finalUrl = `${config.GRAPH_API_ENDPOINT}${encodedPath}${queryString}`;
      console.error(`Full URL: ${finalUrl}`);
    }
    
    return new Promise((resolve, reject) => {
      const options = {
        method: method,
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      };
      
      const req = https.request(finalUrl, options, (res) => {
        let responseData = '';
        
        res.on('data', (chunk) => {
          responseData += chunk;
        });
        
        res.on('end', () => {
          if (res.statusCode >= 200 && res.statusCode < 300) {
            try {
              responseData = responseData ? responseData : '{}';
              const jsonResponse = JSON.parse(responseData);
              resolve(jsonResponse);
            } catch (error) {
              reject(new Error(`Error parsing API response: ${error.message}`));
            }
          } else if (res.statusCode === 401) {
            // Token expired or invalid
            reject(new Error('UNAUTHORIZED'));
          } else {
            reject(new Error(`API call failed with status ${res.statusCode}: ${responseData}`));
          }
        });
      });
      
      req.on('error', (error) => {
        reject(new Error(`Network error during API call: ${error.message}`));
      });
      
      if (data && (method === 'POST' || method === 'PATCH' || method === 'PUT')) {
        req.write(JSON.stringify(data));
      }
      
      req.end();
    });
  } catch (error) {
    console.error('Error calling Graph API:', error);
    throw error;
  }
}

/**
 * Calls Graph API with pagination support to retrieve all results up to maxCount
 * @param {string} accessToken - The access token for authentication
 * @param {string} method - HTTP method (GET only for pagination)
 * @param {string} path - API endpoint path
 * @param {object} queryParams - Initial query parameters
 * @param {number} maxCount - Maximum number of items to retrieve (0 = all)
 * @returns {Promise<object>} - Combined API response with all items
 */
async function callGraphAPIPaginated(accessToken, method, path, queryParams = {}, maxCount = 0) {
  if (method !== 'GET') {
    throw new Error('Pagination only supports GET requests');
  }

  const allItems = [];
  let nextLink = null;
  let currentUrl = path;
  let currentParams = queryParams;

  try {
    do {
      // Make API call
      const response = await callGraphAPI(accessToken, method, currentUrl, null, currentParams);
      
      // Add items from this page
      if (response.value && Array.isArray(response.value)) {
        allItems.push(...response.value);
        console.error(`Pagination: Retrieved ${response.value.length} items, total so far: ${allItems.length}`);
      }

      // Check if we've reached the desired count
      if (maxCount > 0 && allItems.length >= maxCount) {
        console.error(`Pagination: Reached max count of ${maxCount}, stopping`);
        break;
      }

      // Get next page URL
      nextLink = response['@odata.nextLink'];
      
      if (nextLink) {
        // Pass the full nextLink URL directly to callGraphAPI
        currentUrl = nextLink;
        currentParams = {}; // nextLink already contains all params
        console.error(`Pagination: Following nextLink, ${allItems.length} items so far`);
      }
    } while (nextLink);

    // Trim to exact count if needed
    const finalItems = maxCount > 0 ? allItems.slice(0, maxCount) : allItems;

    console.error(`Pagination complete: Retrieved ${finalItems.length} total items`);
    
    return {
      value: finalItems,
      '@odata.count': finalItems.length
    };
  } catch (error) {
    console.error('Error during pagination:', error);
    throw error;
  }
}

/**
 * Makes a request to the Microsoft Graph API that returns a download URL (302 redirect)
 * Used for OneDrive file downloads which return a pre-authenticated download URL
 * @param {string} accessToken - The access token for authentication
 * @param {string} path - API endpoint path
 * @returns {Promise<string>} - The download URL from the redirect
 */
async function callGraphAPIDownload(accessToken, path) {
  // For test tokens, simulate download
  if (config.USE_TEST_MODE && accessToken.startsWith('test_access_token_')) {
    console.error(`TEST MODE: Simulating download for ${path}`);
    return `https://example.com/download/${Date.now()}`;
  }

  return new Promise((resolve, reject) => {
    const fullUrl = `${config.GRAPH_API_ENDPOINT}${path}`;
    console.error(`Making download request: GET ${fullUrl}`);

    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    };

    const req = https.request(fullUrl, options, (res) => {
      // Graph API returns 302 with Location header containing the download URL
      if (res.statusCode === 302 && res.headers.location) {
        resolve(res.headers.location);
      } else if (res.statusCode >= 200 && res.statusCode < 300) {
        // Some endpoints might return the URL in the body instead
        let responseData = '';
        res.on('data', (chunk) => {
          responseData += chunk;
        });
        res.on('end', () => {
          try {
            const jsonResponse = JSON.parse(responseData);
            if (jsonResponse['@microsoft.graph.downloadUrl']) {
              resolve(jsonResponse['@microsoft.graph.downloadUrl']);
            } else {
              reject(new Error('No download URL found in response'));
            }
          } catch (error) {
            reject(new Error(`Error parsing download response: ${error.message}`));
          }
        });
      } else if (res.statusCode === 401) {
        reject(new Error('UNAUTHORIZED'));
      } else {
        let responseData = '';
        res.on('data', (chunk) => {
          responseData += chunk;
        });
        res.on('end', () => {
          reject(new Error(`Download request failed with status ${res.statusCode}: ${responseData}`));
        });
      }
    });

    req.on('error', (error) => {
      reject(new Error(`Network error during download request: ${error.message}`));
    });

    req.end();
  });
}

module.exports = {
  callGraphAPI,
  callGraphAPIPaginated,
  callGraphAPIDownload
};
