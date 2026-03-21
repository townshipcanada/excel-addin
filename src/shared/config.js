/**
 * Township Canada Excel Add-In - Shared Configuration & API Helpers
 *
 * Single source of truth for API URL, storage keys, and common functions
 * used across custom functions, commands, and the task pane.
 *
 * Requires a trial or paid API key. Trial keys can be obtained at:
 * https://townshipcanada.com/api/try?ref=excel
 */

export var API_BASE_URL = "https://townshipcanada.com/api/integrations/trial";
export var MAX_BATCH_SIZE = 200;
export var TRIAL_URL = "https://townshipcanada.com/api/try?ref=excel";

var STORAGE_KEY_API_KEY = "township_api_key";

/**
 * Get the stored API key (trial or paid).
 */
export function getApiKey() {
  return localStorage.getItem(STORAGE_KEY_API_KEY) || "";
}

/**
 * Save an API key to localStorage.
 */
export function setApiKey(key) {
  localStorage.setItem(STORAGE_KEY_API_KEY, key.trim());
}

/**
 * Remove the stored API key.
 */
export function removeApiKey() {
  localStorage.removeItem(STORAGE_KEY_API_KEY);
}

/**
 * Check if an API key is configured.
 */
export function hasApiKey() {
  return !!getApiKey();
}

/**
 * Build request headers with API key.
 */
export function buildHeaders() {
  var headers = {
    "Content-Type": "application/json"
  };
  var apiKey = getApiKey();
  if (apiKey) {
    headers["X-API-Key"] = apiKey;
  }
  return headers;
}

/**
 * Convert a single legal land description via the API.
 */
export async function apiConvertSingle(query) {
  if (!hasApiKey()) {
    throw new Error("NO_API_KEY");
  }

  var response = await fetch(API_BASE_URL + "/convert", {
    method: "POST",
    headers: buildHeaders(),
    body: JSON.stringify({ query: query })
  });

  if (response.status === 401) {
    throw new Error("INVALID_API_KEY");
  }
  if (response.status === 403) {
    throw new Error("TRIAL_EXPIRED");
  }
  if (response.status === 429) {
    throw new Error("TRIAL_LIMIT_REACHED");
  }
  if (!response.ok) {
    var errorBody = await response.json();
    throw new Error(errorBody.message || "API request failed");
  }

  var body = await response.json();
  return body.data;
}

/**
 * Convert a batch of legal land descriptions via the API.
 */
export async function apiConvertBatch(queries) {
  if (!hasApiKey()) {
    throw new Error("NO_API_KEY");
  }

  var response = await fetch(API_BASE_URL + "/convert-batch", {
    method: "POST",
    headers: buildHeaders(),
    body: JSON.stringify({ batch: queries })
  });

  if (response.status === 401) {
    throw new Error("INVALID_API_KEY");
  }
  if (response.status === 403) {
    throw new Error("TRIAL_EXPIRED");
  }
  if (response.status === 429) {
    throw new Error("TRIAL_LIMIT_REACHED");
  }
  if (!response.ok) {
    var errorBody = await response.json();
    throw new Error(errorBody.message || "Batch API request failed");
  }

  return await response.json();
}

/**
 * Get current usage information for the connected API key.
 */
export async function apiGetUsage() {
  if (!hasApiKey()) {
    return { plan: "none", apiKeyValid: false };
  }

  try {
    var response = await fetch(API_BASE_URL + "/usage", {
      method: "GET",
      headers: buildHeaders()
    });

    if (!response.ok) {
      return { plan: "none", apiKeyValid: false };
    }

    return (await response.json()).data;
  } catch (e) {
    return { plan: "none", apiKeyValid: false };
  }
}
