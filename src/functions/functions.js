/**
 * Township Canada Excel Add-In - Custom Functions
 *
 * Registered under the TOWNSHIP_CANADA namespace:
 *   =TOWNSHIP_CANADA.CONVERT("NW-25-24-1-W5")    -> "52.123456, -114.654321"
 *   =TOWNSHIP_CANADA.LAT("NW-25-24-1-W5")        -> 52.123456
 *   =TOWNSHIP_CANADA.LNG("NW-25-24-1-W5")        -> -114.654321
 *   =TOWNSHIP_CANADA.PROVINCE("NW-25-24-1-W5")   -> "Alberta"
 *
 * 100 sample locations work without an API key.
 * For full access, get a key at: townshipcanada.com/api
 */

import { apiConvertSingle } from "../shared/config";

/**
 * Shared error handler for all custom functions.
 * Returns a user-friendly string for known API error codes.
 */
function handleFunctionError(e) {
  switch (e.message) {
    case "NO_API_KEY":
      return "API key required";
    case "INVALID_API_KEY":
      return "Invalid API key";
    case "RATE_LIMIT_REACHED":
      return "Rate limit reached";
    default:
      return "Error: " + e.message;
  }
}

/**
 * Convert a Canadian legal land description to GPS coordinates.
 * Returns "latitude, longitude" as a string.
 *
 * Supports DLS (AB, SK, MB), NTS (BC), Geographic Townships (ON),
 * River Lots, UWI, and FPS Grid formats.
 *
 * @customfunction
 * @param {string} lld The legal land description (e.g., "NW-25-24-1-W5").
 * @returns {string} GPS coordinates as "latitude, longitude".
 */
async function convert(lld) {
  if (!lld || typeof lld !== "string" || !lld.trim()) {
    return "";
  }

  try {
    var result = await apiConvertSingle(lld.trim());
    if (result.latitude !== null && result.longitude !== null) {
      return result.latitude.toFixed(6) + ", " + result.longitude.toFixed(6);
    }
    return "Not found";
  } catch (e) {
    return handleFunctionError(e);
  }
}

/**
 * Convert a legal land description and return only the latitude.
 *
 * @customfunction
 * @param {string} lld The legal land description (e.g., "NW-25-24-1-W5").
 * @returns {number} Latitude in decimal degrees.
 */
async function lat(lld) {
  if (!lld || typeof lld !== "string" || !lld.trim()) {
    return "";
  }

  try {
    var result = await apiConvertSingle(lld.trim());
    if (result.latitude !== null) {
      return result.latitude;
    }
    return "Not found";
  } catch (e) {
    return handleFunctionError(e);
  }
}

/**
 * Convert a legal land description and return only the longitude.
 *
 * @customfunction
 * @param {string} lld The legal land description (e.g., "NW-25-24-1-W5").
 * @returns {number} Longitude in decimal degrees.
 */
async function lng(lld) {
  if (!lld || typeof lld !== "string" || !lld.trim()) {
    return "";
  }

  try {
    var result = await apiConvertSingle(lld.trim());
    if (result.longitude !== null) {
      return result.longitude;
    }
    return "Not found";
  } catch (e) {
    return handleFunctionError(e);
  }
}

/**
 * Convert a legal land description and return the province name.
 *
 * @customfunction
 * @param {string} lld The legal land description (e.g., "NW-25-24-1-W5").
 * @returns {string} Province name (e.g., "Alberta").
 */
async function province(lld) {
  if (!lld || typeof lld !== "string" || !lld.trim()) {
    return "";
  }

  try {
    var result = await apiConvertSingle(lld.trim());
    if (result.province) {
      return result.province;
    }
    return "Not found";
  } catch (e) {
    return handleFunctionError(e);
  }
}

// Register custom functions with Office
CustomFunctions.associate("CONVERT", convert);
CustomFunctions.associate("LAT", lat);
CustomFunctions.associate("LNG", lng);
CustomFunctions.associate("PROVINCE", province);
