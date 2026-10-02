const {EventEmitter} = require('events');

const iterateStream = verquire('utils/iterate-stream');

describe('iterateStream', () => {
  it('does not miss data emitted during resume', async () => {
    class SynchronousResumeStream extends EventEmitter {
      resume() {
        if (!this.resumed) {
          this.resumed = true;
          this.emit('data', 'chunk');
          setTimeout(() => this.emit('end'), 20);
        }
      }

      pause() {}
    }

    const iterator = iterateStream(new SynchronousResumeStream());
    let settled = false;
    const next = iterator.next().then(result => {
      settled = true;
      return result;
    });

    await new Promise(resolve => {
      setImmediate(resolve);
    });
    expect(settled).to.equal(true);
    expect(await next).to.deep.equal({value: 'chunk', done: false});
    expect(await iterator.next()).to.deep.equal({value: undefined, done: true});
  });
});