/**
 * Backward-compatible alias — the regular onedrive-upload handler now
 * picks simple vs chunked upload automatically based on buffer size, so
 * this is just a passthrough. Kept so any older callers of
 * `onedrive-upload-large` still work.
 */
module.exports = require('./upload');
