/**
 * Township Canada Excel Add-In - Custom Functions
 *
 * Registered under the TOWNSHIP namespace:
 *   =TOWNSHIP.CONVERT("NW-25-24-1-W5")           -> "52.123456, -114.654321"
 *   =TOWNSHIP.LAT("NW-25-24-1-W5")               -> 52.123456
 *   =TOWNSHIP.LNG("NW-25-24-1-W5")               -> -114.654321
 *   =TOWNSHIP.PROVINCE("NW-25-24-1-W5")          -> "Alberta"
 *
 * Requires a Township Canada API key (trial or paid).
 * Get a free trial key at: townshipcanada.com/api/try
 */

import { apiConvertSingle } from "../shared/config";

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
    if (e.message === "NO_API_KEY") {
      return "API key required";
    }
    if (e.message === "TRIAL_EXPIRED") {
      return "Trial expired";
    }
    if (e.message === "TRIAL_LIMIT_REACHED") {
      return "Trial limit reached";
    }
    if (e.message === "INVALID_API_KEY") {
      return "Invalid API key";
    }
    return "Error: " + e.message;
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
    if (e.message === "NO_API_KEY") {
      return "API key required";
    }
    if (e.message === "TRIAL_EXPIRED" || e.message === "TRIAL_LIMIT_REACHED") {
      return "Trial ended";
    }
    if (e.message === "INVALID_API_KEY") {
      return "Invalid API key";
    }
    return "Error: " + e.message;
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
    if (e.message === "NO_API_KEY") {
      return "API key required";
    }
    if (e.message === "TRIAL_EXPIRED" || e.message === "TRIAL_LIMIT_REACHED") {
      return "Trial ended";
    }
    if (e.message === "INVALID_API_KEY") {
      return "Invalid API key";
    }
    return "Error: " + e.message;
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
    if (e.message === "NO_API_KEY") {
      return "API key required";
    }
    if (e.message === "TRIAL_EXPIRED" || e.message === "TRIAL_LIMIT_REACHED") {
      return "Trial ended";
    }
    if (e.message === "INVALID_API_KEY") {
      return "Invalid API key";
    }
    return "Error: " + e.message;
  }
}

// Register custom functions with Office
CustomFunctions.associate("CONVERT", convert);
CustomFunctions.associate("LAT", lat);
CustomFunctions.associate("LNG", lng);
CustomFunctions.associate("PROVINCE", province);
