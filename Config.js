/* ---------- CONFIGURATION ---------- */
// This file contains all the global variables and settings for the application.

// Live Source Sheet IDs (for production environment)
var LIVE_SOURCE_IDS = [
  '1URF6p99IPY8_kPCNOhYigL_S7UMvyQAdCyB1dCtVGrY', //MDS Unity (New)
  '1bgOimFOdjXcFZqWzWVvnUw1xRnyhWCAdGrxy1JSlVR4', //YEST
  '1eXJzHNsitv86igh8oAg-HAmz3LXqT9A-aCz9pl3zEvo', //RCC
  '1YscvMF-vB7HgfzYrVCWDCIxGAuerWyfcG2_jZgDttVA', //Wave
  '1NzoxehLtg55gXgBEyZTwGxdl2wltioGnv9Dn_7-HwhI', //TBM - ADL
  '10WVTn2rtFjLJzN4943Nk4G8JsOm2qt9Sa1S6-aDU6ZM', //MDS ADL (New)
  '1hKtLqUgaGhHZvcFm_BydLBYAPUJ9IcyYtDhZGezy2u8', //MDS EVEREST
  '1D207Z49WrW7UgEvbvCRvqB_uVP1IEu1lB89Vxkqt1gY', //MDS Pathway (New)
  '19PfI2wDAUB_qKFHkMJzUHIkLSwzaNmS91h63P1aHrvw', //MDS Star (New)
  '1jlSPG05hqVh-RBsDzFmdYT_e2eA2cMEz9A4quaeKG8I', //TBM EVEREST
  '1XDZidmShTfyku6rWH5YYJBWYTdAL3RlnqZyByqGYRXo', //TBM District
  '1jodXK36Y7ojU7g9iij6FAZCqLS37BTY407Hgv00SWAk', //TBM PATH (New)
  '1ATCASWpQ3Ju5481B-Uh_1fhIDfAafnxnKd1903K5jA0', //TBM - STARLABS
  '1ZX36vmCYR0D5_8Y-fx9sLct46As_bMD6mGIxm9aBPN8'  //TBM UNITY
];

// Development Source Sheet IDs (for testing)
var SOURCE_IDS = [
  '1QgLFeJiw8vuF849rGokWEDDsKyIpd5iKY9OY_SdrA4Q', //DEV RCC
  '14NZWsm3HVlEDVwV-ncWnZRF3GK3hKByen0euSbd6CUg', //DEV TBM
];

// Master Spreadsheet IDs
var LIVE_SPREADSHEET_ID = '161uw5s1lOwhV7YTX8uLKDw6TMl_QLEpSqxmHEftZzVA'; // Production Master Sheet
var MASTER_SPREADSHEET_ID = '14I2UGjK3Vmsbya9PJZY16SxIDKotply3ysLHkWI-Hpw'; // Development Master Sheet

// --- !! IMPORTANT !! ---
// After deploying the script as a Web App, paste the correct URL here.
var WEB_APP_URL = 'https://script.google.com/a/macros/bhg.llc/s/AKfycby-X9fXPxjzelSuWxORy8dDAcXOEQ9yf_Aqy3lsSw/exec'; 

// Sheet Names
var DEST_SHEET_NAME = 'Master';
var LOG_SHEET_NAME = 'Log';
var DUPLICATES_SHEET_NAME = 'Duplicates';
var CHECKER_SHEET_NAME = 'Checker';
var HELPER_SHEET_NAME = '_QueryHelper';

// Data Validation & Formatting
var MANDATORY_FIELDS = ['DATE', 'FIRST NAME', 'LAST NAME', 'DOB'];
var NEW_ROW_DATE_HIGHLIGHT = '#ffcccc';

// Script Properties Prefixes (for storing state)
var PROP_PREFIX = 'lastRow_';
var CONS_PROP_PREFIX = 'consolidate_lastRow_';
var DUP_GROUP_PREFIX = 'dup_logged_';

// Log Headers
var LOG_HEADERS = ['TIMESTAMP', 'REASON', 'SOURCE_SPREADSHEET_ID', 'SOURCE_SHEET_NAME', 'ROW_NUMBER', 'ORIGINAL_VALUES', 'LINK_TO_ROW', 'DUPLICATE_LINKS', 'TRIGGERING_USER'];
/* ------------------------------- */