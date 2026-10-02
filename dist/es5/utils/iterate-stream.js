"use strict";

module.exports = async function* iterateStream(stream) {
  const isNode = typeof process !== 'undefined' && process.versions && process.versions.node;
  if (isNode && stream.readableObjectMode && stream[Symbol.asyncIterator]) {
    for await (const data of stream) {
      yield data;
    }
    return;
  }
  const contents = [];
  stream.on('data', data => contents.push(data));
  let resolveStreamEndedPromise;
  const streamEndedPromise = new Promise(resolve => {
    resolveStreamEndedPromise = resolve;
  });
  let ended = false;
  stream.on('end', () => {
    ended = true;
    resolveStreamEndedPromise();
  });
  let error = false;
  stream.on('error', err => {
    error = err;
    resolveStreamEndedPromise();
  });
  while (!ended || contents.length > 0) {
    if (contents.length === 0) {
      const dataPromise = once(stream, 'data');
      stream.resume();
      // eslint-disable-next-line no-await-in-loop
      await Promise.race([dataPromise, streamEndedPromise]);
    } else {
      stream.pause();
      const data = contents.shift();
      yield data;
    }
    if (error) throw error;
  }
  resolveStreamEndedPromise();
};
function once(eventEmitter, type) {
  // TODO: Use require('events').once when node v10 is dropped
  return new Promise(resolve => {
    let fired = false;
    const handler = () => {
      if (!fired) {
        fired = true;
        eventEmitter.removeListener(type, handler);
        resolve();
      }
    };
    eventEmitter.addListener(type, handler);
  });
}
//# sourceMappingURL=iterate-stream.js.map
