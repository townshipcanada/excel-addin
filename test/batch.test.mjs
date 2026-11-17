// Tests for the Batch API client in src/shared/config.js.
// Run with: npm test (Node's built-in test runner, no extra dependencies).

import { test, beforeEach, afterEach } from "node:test";
import assert from "node:assert/strict";

// config.js reads the API key from localStorage on every call.
var store = {};
globalThis.localStorage = {
  getItem: function (k) {
    return Object.prototype.hasOwnProperty.call(store, k) ? store[k] : null;
  },
  setItem: function (k, v) {
    store[k] = String(v);
  },
  removeItem: function (k) {
    delete store[k];
  }
};

var config = await import("../src/shared/config.js");
var sample = await import("../src/shared/sampleData.js");

var realFetch = globalThis.fetch;
var calls;

// A FeatureCollection shaped like api/services/batch/handler.js builds it.
function featureCollection(term, i) {
  var props = { search_term: term, legal_location: term, province: "Alberta" };
  return {
    type: "FeatureCollection",
    features: [
      { type: "Feature", geometry: { type: "Polygon", coordinates: [] }, properties: Object.assign({ shape: "grid" }, props) },
      { type: "Feature", geometry: { type: "Point", coordinates: [-110 - i / 1000, 50 + i / 1000] }, properties: Object.assign({ shape: "centroid" }, props) }
    ]
  };
}

// Mimics the Batch API: bare JSON array in, 400 for non-arrays, 413 above
// 100 items, a bare array of FeatureCollections out, unmatched inputs omitted.
function mockBatchApi(unmatched) {
  globalThis.fetch = async function (url, options) {
    var body = JSON.parse(options.body);
    calls.push({ url: url, body: body });
    if (!Array.isArray(body)) {
      return jsonResponse(400, { error: "Bad Request. Input must be an array.", message: "The request body must be a JSON array." });
    }
    if (body.length > 100) {
      return jsonResponse(413, { error: "Payload Too Large.", message: "Maximum payload size is 100 items." });
    }
    var out = [];
    body.forEach(function (term) {
      if (unmatched.indexOf(term) === -1) {
        out.push(featureCollection(term.trim(), Number(term.split("-")[1])));
      }
    });
    return jsonResponse(200, out);
  };
}

function jsonResponse(status, body) {
  return {
    status: status,
    ok: status >= 200 && status < 300,
    json: async function () {
      return body;
    }
  };
}

// Descriptions that are not in the offline sample data, so they go to the API.
function apiTerms(n) {
  var terms = [];
  for (var i = 1; i <= n; i++) {
    var term = "NE-" + i + "-77-9-W5";
    assert.equal(sample.hasSampleData(term), false);
    terms.push(term);
  }
  return terms;
}

beforeEach(function () {
  calls = [];
  store = { township_api_key: "test-key" };
});

afterEach(function () {
  globalThis.fetch = realFetch;
});

test("MAX_BATCH_SIZE matches the Batch API limit of 100", function () {
  assert.equal(config.MAX_BATCH_SIZE, 100);
});

test("sends the request body as a bare JSON array", async function () {
  mockBatchApi([]);
  var queries = apiTerms(3);
  await config.apiConvertBatch(queries);
  assert.equal(calls.length, 1);
  assert.deepEqual(calls[0].body, queries);
});

test("splits more than 100 descriptions into requests of at most 100", async function () {
  mockBatchApi([]);
  var queries = apiTerms(250);
  var res = await config.apiConvertBatch(queries);
  assert.deepEqual(
    calls.map(function (c) {
      return c.body.length;
    }),
    [100, 100, 50]
  );
  assert.equal(res.data.length, 250);
  assert.equal(res.data[249].legal_location, "NE-250-77-9-W5");
});

test("aligns results to inputs by search_term when unmatched inputs are omitted", async function () {
  var queries = apiTerms(5);
  mockBatchApi([queries[1], queries[3]]);
  var res = await config.apiConvertBatch(queries);
  assert.equal(res.data.length, 5);
  assert.deepEqual(
    res.data.map(function (r) {
      return r.legal_location;
    }),
    [queries[0], null, queries[2], null, queries[4]]
  );
  assert.equal(res.data[1].latitude, null);
  assert.equal(res.data[3].latitude, null);
  assert.equal(res.data[4].latitude, 50.005);
  assert.equal(res.data[4].longitude, -110.005);
  assert.equal(res.data[4].province, "Alberta");
});

test("matches inputs with surrounding spaces or quotes, as the API strips them", async function () {
  mockBatchApi([]);
  globalThis.fetch = (function (inner) {
    return async function (url, options) {
      var res = await inner(url, options);
      var body = await res.json();
      body.forEach(function (fc) {
        fc.features.forEach(function (f) {
          f.properties.search_term = f.properties.search_term.replace(/['"]/g, "");
        });
      });
      return jsonResponse(res.status, body);
    };
  })(globalThis.fetch);
  var res = await config.apiConvertBatch([" 'NE-7-77-9-W5' "]);
  assert.equal(res.data[0].latitude, 50.007);
  assert.equal(res.data[0].longitude, -110.007);
});

test("keeps offline sample hits in place alongside API results", async function () {
  var queries = apiTerms(2);
  mockBatchApi([queries[1]]);
  var res = await config.apiConvertBatch([queries[0], "NE-1-25-1-W4", queries[1]]);
  assert.deepEqual(calls[0].body, queries);
  assert.equal(res.data[0].legal_location, queries[0]);
  assert.equal(res.data[1].latitude, 50.6389);
  assert.equal(res.data[1].legal_location, "NE-1-25-1-W4");
  assert.equal(res.data[2].latitude, null);
});

test("rejects a response that is not an array", async function () {
  globalThis.fetch = async function () {
    return jsonResponse(200, { data: [] });
  };
  await assert.rejects(config.apiConvertBatch(apiTerms(1)), /Unexpected Batch API response/);
});
