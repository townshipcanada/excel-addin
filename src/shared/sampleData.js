/**
 * Township Canada Excel Add-In - Sample Data
 *
 * 100 hardcoded DLS values with GPS coordinates for offline demo.
 * Lets users try the add-in instantly without an API key,
 * encouraging them to purchase a batch key for full access.
 *
 * Coverage: Alberta (W4/W5/W6), Saskatchewan (W2/W3), Manitoba (W1)
 * Formats: Quarter-section (e.g., NW-25-24-1-W5) and LSD (e.g., 10-25-24-1-W5)
 */

var SAMPLE_DATA = {
  // ── Alberta W4 (Quarter-section) ──────────────────────────────
  "NE-1-25-1-W4":   { latitude: 50.63890, longitude: -110.00540, province: "Alberta" },
  "NW-1-25-1-W4":   { latitude: 50.63890, longitude: -110.02160, province: "Alberta" },
  "SE-1-25-1-W4":   { latitude: 50.62460, longitude: -110.00540, province: "Alberta" },
  "SW-1-25-1-W4":   { latitude: 50.62460, longitude: -110.02160, province: "Alberta" },
  "NE-12-30-15-W4":  { latitude: 51.28370, longitude: -112.03680, province: "Alberta" },
  "NW-12-30-15-W4":  { latitude: 51.28370, longitude: -112.05300, province: "Alberta" },
  "SE-36-42-28-W4":  { latitude: 52.56460, longitude: -113.85780, province: "Alberta" },
  "NW-6-50-10-W4":   { latitude: 53.35790, longitude: -111.42580, province: "Alberta" },
  "NE-15-60-20-W4":  { latitude: 54.28460, longitude: -112.77180, province: "Alberta" },
  "SW-22-70-5-W4":   { latitude: 55.20630, longitude: -110.72780, province: "Alberta" },

  // ── Alberta W4 (LSD) ─────────────────────────────────────────
  "10-1-25-1-W4":   { latitude: 50.63530, longitude: -110.01350, province: "Alberta" },
  "01-1-25-1-W4":   { latitude: 50.62820, longitude: -110.02160, province: "Alberta" },
  "16-1-25-1-W4":   { latitude: 50.64250, longitude: -110.01350, province: "Alberta" },
  "06-12-30-15-W4":  { latitude: 50.27650, longitude: -112.05300, province: "Alberta" },
  "13-36-42-28-W4":  { latitude: 52.56820, longitude: -113.85780, province: "Alberta" },
  "11-6-50-10-W4":   { latitude: 53.35430, longitude: -111.42580, province: "Alberta" },
  "09-15-60-20-W4":  { latitude: 54.28100, longitude: -112.77180, province: "Alberta" },
  "04-22-70-5-W4":   { latitude: 55.19910, longitude: -110.72780, province: "Alberta" },

  // ── Alberta W5 (Quarter-section) ──────────────────────────────
  "NW-25-24-1-W5":  { latitude: 50.56750, longitude: -114.06320, province: "Alberta" },
  "NE-25-24-1-W5":  { latitude: 50.56750, longitude: -114.04700, province: "Alberta" },
  "SE-25-24-1-W5":  { latitude: 50.55320, longitude: -114.04700, province: "Alberta" },
  "SW-25-24-1-W5":  { latitude: 50.55320, longitude: -114.06320, province: "Alberta" },
  "NE-1-26-2-W5":   { latitude: 50.72560, longitude: -114.20100, province: "Alberta" },
  "NW-1-26-2-W5":   { latitude: 50.72560, longitude: -114.21720, province: "Alberta" },
  "SE-10-30-5-W5":  { latitude: 51.22880, longitude: -114.67540, province: "Alberta" },
  "NW-18-35-8-W5":  { latitude: 51.70120, longitude: -115.14960, province: "Alberta" },
  "NE-22-40-3-W5":  { latitude: 52.14270, longitude: -114.40940, province: "Alberta" },
  "SW-30-45-10-W5": { latitude: 52.62140, longitude: -115.41860, province: "Alberta" },
  "NE-5-50-5-W5":   { latitude: 53.32560, longitude: -114.67540, province: "Alberta" },
  "NW-15-55-2-W5":  { latitude: 53.78710, longitude: -114.28070, province: "Alberta" },
  "SE-20-60-7-W5":  { latitude: 54.24530, longitude: -114.94160, province: "Alberta" },
  "NE-33-65-12-W5": { latitude: 54.72430, longitude: -115.68760, province: "Alberta" },
  "SW-8-70-1-W5":   { latitude: 55.14750, longitude: -114.12280, province: "Alberta" },
  "NW-14-75-4-W5":  { latitude: 55.60740, longitude: -114.52940, province: "Alberta" },

  // ── Alberta W5 (LSD) ─────────────────────────────────────────
  "10-25-24-1-W5":  { latitude: 50.56390, longitude: -114.05510, province: "Alberta" },
  "01-25-24-1-W5":  { latitude: 50.55680, longitude: -114.06320, province: "Alberta" },
  "16-25-24-1-W5":  { latitude: 50.57110, longitude: -114.05510, province: "Alberta" },
  "08-1-26-2-W5":   { latitude: 50.72200, longitude: -114.20910, province: "Alberta" },
  "05-10-30-5-W5":  { latitude: 51.22160, longitude: -114.67540, province: "Alberta" },
  "14-18-35-8-W5":  { latitude: 51.70480, longitude: -115.14960, province: "Alberta" },
  "11-22-40-3-W5":  { latitude: 52.13910, longitude: -114.40940, province: "Alberta" },
  "03-30-45-10-W5": { latitude: 52.61420, longitude: -115.41860, province: "Alberta" },
  "12-5-50-5-W5":   { latitude: 53.32920, longitude: -114.67540, province: "Alberta" },
  "09-15-55-2-W5":  { latitude: 53.78350, longitude: -114.28070, province: "Alberta" },
  "07-20-60-7-W5":  { latitude: 54.24170, longitude: -114.94160, province: "Alberta" },
  "15-33-65-12-W5": { latitude: 54.72790, longitude: -115.68760, province: "Alberta" },

  // ── Alberta W6 (Quarter-section) ──────────────────────────────
  "NE-1-55-1-W6":   { latitude: 53.74170, longitude: -118.01520, province: "Alberta" },
  "NW-1-55-1-W6":   { latitude: 53.74170, longitude: -118.03140, province: "Alberta" },
  "SE-10-58-3-W6":  { latitude: 54.01780, longitude: -118.31540, province: "Alberta" },
  "NW-20-60-5-W6":  { latitude: 54.24530, longitude: -118.63560, province: "Alberta" },
  "NE-35-62-2-W6":  { latitude: 54.47580, longitude: -118.21940, province: "Alberta" },
  "SW-15-65-4-W6":  { latitude: 54.69880, longitude: -118.49540, province: "Alberta" },

  // ── Alberta W6 (LSD) ─────────────────────────────────────────
  "10-1-55-1-W6":   { latitude: 53.73810, longitude: -118.02330, province: "Alberta" },
  "05-10-58-3-W6":  { latitude: 54.01060, longitude: -118.31540, province: "Alberta" },
  "14-20-60-5-W6":  { latitude: 54.24890, longitude: -118.63560, province: "Alberta" },
  "08-35-62-2-W6":  { latitude: 54.47220, longitude: -118.21940, province: "Alberta" },

  // ── Saskatchewan W2 (Quarter-section) ─────────────────────────
  "NE-1-20-1-W2":   { latitude: 50.28370, longitude: -102.00540, province: "Saskatchewan" },
  "NW-1-20-1-W2":   { latitude: 50.28370, longitude: -102.02160, province: "Saskatchewan" },
  "SE-15-25-5-W2":  { latitude: 50.63890, longitude: -102.68540, province: "Saskatchewan" },
  "NW-22-30-10-W2": { latitude: 51.28370, longitude: -103.36540, province: "Saskatchewan" },
  "NE-36-35-15-W2": { latitude: 51.74520, longitude: -104.04540, province: "Saskatchewan" },
  "SW-10-40-20-W2": { latitude: 52.10060, longitude: -104.72540, province: "Saskatchewan" },
  "NE-5-45-8-W2":   { latitude: 52.55640, longitude: -103.09340, province: "Saskatchewan" },
  "NW-28-50-3-W2":  { latitude: 53.01920, longitude: -102.41340, province: "Saskatchewan" },

  // ── Saskatchewan W2 (LSD) ────────────────────────────────────
  "10-1-20-1-W2":   { latitude: 50.28010, longitude: -102.01350, province: "Saskatchewan" },
  "06-15-25-5-W2":  { latitude: 50.63170, longitude: -102.68540, province: "Saskatchewan" },
  "13-22-30-10-W2": { latitude: 51.28730, longitude: -103.36540, province: "Saskatchewan" },
  "09-36-35-15-W2": { latitude: 51.74160, longitude: -104.04540, province: "Saskatchewan" },
  "04-10-40-20-W2": { latitude: 52.09340, longitude: -104.72540, province: "Saskatchewan" },

  // ── Saskatchewan W3 (Quarter-section) ─────────────────────────
  "NE-1-20-1-W3":   { latitude: 50.28370, longitude: -106.00540, province: "Saskatchewan" },
  "NW-1-20-1-W3":   { latitude: 50.28370, longitude: -106.02160, province: "Saskatchewan" },
  "SE-12-25-5-W3":  { latitude: 50.63890, longitude: -106.64940, province: "Saskatchewan" },
  "NW-18-30-10-W3": { latitude: 51.24600, longitude: -107.38140, province: "Saskatchewan" },
  "NE-25-35-8-W3":  { latitude: 51.74520, longitude: -107.06140, province: "Saskatchewan" },
  "SW-6-40-15-W3":  { latitude: 52.05600, longitude: -108.01340, province: "Saskatchewan" },
  "NE-30-45-3-W3":  { latitude: 52.62140, longitude: -106.37340, province: "Saskatchewan" },
  "NW-15-50-12-W3": { latitude: 53.01920, longitude: -107.69340, province: "Saskatchewan" },
  "SE-20-55-6-W3":  { latitude: 53.47460, longitude: -106.81340, province: "Saskatchewan" },
  "NE-33-60-1-W3":  { latitude: 54.00210, longitude: -106.05340, province: "Saskatchewan" },

  // ── Saskatchewan W3 (LSD) ────────────────────────────────────
  "10-1-20-1-W3":   { latitude: 50.28010, longitude: -106.01350, province: "Saskatchewan" },
  "06-12-25-5-W3":  { latitude: 50.63170, longitude: -106.64940, province: "Saskatchewan" },
  "14-18-30-10-W3": { latitude: 51.24960, longitude: -107.38140, province: "Saskatchewan" },
  "09-25-35-8-W3":  { latitude: 51.74160, longitude: -107.06140, province: "Saskatchewan" },
  "03-6-40-15-W3":  { latitude: 52.04880, longitude: -108.01340, province: "Saskatchewan" },

  // ── Manitoba W1 (Quarter-section) ─────────────────────────────
  "NE-1-10-1-W1":   { latitude: 49.39370, longitude: -97.36540, province: "Manitoba" },
  "NW-1-10-1-W1":   { latitude: 49.39370, longitude: -97.38160, province: "Manitoba" },
  "SE-1-10-1-W1":   { latitude: 49.37940, longitude: -97.36540, province: "Manitoba" },
  "SW-1-10-1-W1":   { latitude: 49.37940, longitude: -97.38160, province: "Manitoba" },
  "NE-15-15-5-W1":  { latitude: 49.83750, longitude: -98.00540, province: "Manitoba" },
  "NW-22-20-10-W1": { latitude: 50.28370, longitude: -98.72540, province: "Manitoba" },
  "SE-30-25-3-W1":  { latitude: 50.60660, longitude: -97.68340, province: "Manitoba" },
  "NE-6-30-8-W1":   { latitude: 51.16530, longitude: -98.44540, province: "Manitoba" },
  "SW-18-35-12-W1": { latitude: 51.60070, longitude: -99.01340, province: "Manitoba" },
  "NE-25-40-5-W1":  { latitude: 52.14270, longitude: -98.00540, province: "Manitoba" },

  // ── Manitoba W1 (LSD) ────────────────────────────────────────
  "10-1-10-1-W1":   { latitude: 49.39010, longitude: -97.37350, province: "Manitoba" },
  "01-1-10-1-W1":   { latitude: 49.38300, longitude: -97.38160, province: "Manitoba" },
  "16-1-10-1-W1":   { latitude: 49.39730, longitude: -97.37350, province: "Manitoba" },
  "06-15-15-5-W1":  { latitude: 49.83030, longitude: -98.00540, province: "Manitoba" },
  "13-22-20-10-W1": { latitude: 50.28730, longitude: -98.72540, province: "Manitoba" },
  "09-30-25-3-W1":  { latitude: 50.60300, longitude: -97.68340, province: "Manitoba" },
  "04-6-30-8-W1":   { latitude: 51.15810, longitude: -98.44540, province: "Manitoba" },
  "11-18-35-12-W1": { latitude: 51.60430, longitude: -99.01340, province: "Manitoba" }
};

/**
 * Look up a legal land description in the sample data.
 * Returns the matching record or null if not found.
 *
 * Normalizes input by uppercasing and trimming whitespace.
 *
 * @param {string} query - The legal land description to look up.
 * @returns {{ latitude: number, longitude: number, province: string } | null}
 */
export function lookupSampleData(query) {
  if (!query) return null;
  var normalized = query.trim().toUpperCase();
  return SAMPLE_DATA[normalized] || null;
}

/**
 * Check if a query exists in the sample data.
 *
 * @param {string} query - The legal land description to check.
 * @returns {boolean}
 */
export function hasSampleData(query) {
  return lookupSampleData(query) !== null;
}
