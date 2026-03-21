/**
 * Township Canada Excel Add-In - Shared Configuration & API Helpers
 *
 * Single source of truth for API URL, storage keys, and common functions
 * used across custom functions, commands, and the task pane.
 *
 * Requires a trial or paid API key. Trial keys can be obtained at:
 * https://townshipcanada.com/api/try?ref=excel
 *
 * Both trial and paid keys use the same API contract (GeoJSON).
 * Trial keys (tc_trial_...) route to the integration trial endpoint;
 * paid keys route to developer.townshipcanada.com.
 */

export var TRIAL_API_BASE_URL = "https://townshipcanada.com/api/integrations/trial";
export var PAID_API_BASE_URL = "https://developer.townshipcanada.com";
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
 * Check if the stored API key is a trial key.
 * Trial keys use the prefix "tc_trial_".
 */
export function isTrialKey() {
  return getApiKey().indexOf("tc_trial_") === 0;
}

/**
 * Get the appropriate API base URL based on the key type.
 */
export function getApiBaseUrl() {
  return isTrialKey() ? TRIAL_API_BASE_URL : PAID_API_BASE_URL;
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
 * Extract coordinates from a GeoJSON FeatureCollection response.
 * Finds the centroid feature and returns a flat object.
 */
function extractFromFeatureCollection(fc) {
  var features = fc.features || [];
  var centroid = null;
  for (var i = 0; i < features.length; i++) {
    if (features[i].properties && features[i].properties.shape === "centroid") {
      centroid = features[i];
      break;
    }
  }
  if (!centroid) {
    centroid = features[0];
  }
  if (!centroid) {
    return { latitude: null, longitude: null, legal_location: null, province: null };
  }
  return {
    latitude: centroid.geometry.coordinates[1],
    longitude: centroid.geometry.coordinates[0],
    legal_location: centroid.properties.legal_location || "",
    province: centroid.properties.province || ""
  };
}

/**
 * Convert a single legal land description via the API.
 * Uses GET /search/legal-location — same contract for trial and paid keys.
 */
export async function apiConvertSingle(query) {
  if (!hasApiKey()) {
    throw new Error("NO_API_KEY");
  }

  var response = await fetch(getApiBaseUrl() + "/search/legal-location?location=" + encodeURIComponent(query), {
    method: "GET",
    headers: buildHeaders()
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
  return extractFromFeatureCollection(body);
}

/**
 * Convert a batch of legal land descriptions via the API.
 * Uses POST /batch/legal-location — same contract for trial and paid keys.
 */
export async function apiConvertBatch(queries) {
  if (!hasApiKey()) {
    throw new Error("NO_API_KEY");
  }

  var response = await fetch(getApiBaseUrl() + "/batch/legal-location", {
    method: "POST",
    headers: buildHeaders(),
    body: JSON.stringify(queries)
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
 * Usage endpoint is only available for trial keys.
 */
export async function apiGetUsage() {
  if (!hasApiKey()) {
    return { plan: "none", apiKeyValid: false };
  }

  if (!isTrialKey()) {
    return { plan: "api_key", apiKeyValid: true };
  }

  try {
    var response = await fetch(TRIAL_API_BASE_URL + "/usage", {
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
