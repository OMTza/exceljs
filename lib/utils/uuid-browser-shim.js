'use strict';

// Minimal RFC4122 v4 UUID generator, used only in browser builds in place of the
// real `uuid` package, whose modern syntax (`??=`) breaks browserify's parser.
function v4() {
  const bytes = new Uint8Array(16);
  // eslint-disable-next-line no-undef
  crypto.getRandomValues(bytes);
  // eslint-disable-next-line no-bitwise
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  // eslint-disable-next-line no-bitwise
  bytes[8] = (bytes[8] & 0x3f) | 0x80;
  const hex = Array.from(bytes, byte => byte.toString(16).padStart(2, '0'));
  return [
    hex.slice(0, 4).join(''),
    hex.slice(4, 6).join(''),
    hex.slice(6, 8).join(''),
    hex.slice(8, 10).join(''),
    hex.slice(10, 16).join(''),
  ].join('-');
}

module.exports = {v4};
