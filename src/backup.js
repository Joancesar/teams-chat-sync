const fs = require('fs');
const path = require('path');
const util = require('util');
const axios = require('axios');

const fsAPI = {
  writeFile: util.promisify(fs.writeFile),
  open: util.promisify(fs.open),
  write: util.promisify(fs.write),
  close: util.promisify(fs.close),
  readdir: util.promisify(fs.readdir),
  readFile: util.promisify(fs.readFile)
};

const FILENAME_MATCH = /^(messages|delta)-.*\.json$/;
const UPLOADED_IMAGE_MATCH = /https:\/\/graph.microsoft.com\/beta\/chats([^"]*)/g;

class Backup {
  constructor({ chatId, authToken, target }) {
    this.target = target;
    this.chatId = chatId;
    this.instance = axios.create({
      headers: {
        Accept: 'application/json, text/plain, */*',
        Authorization: `Bearer ${authToken}`,
        'Sec-Fetch-Mode': 'cors',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.120 Safari/537.36'
      }
    });
  }

  async run() {
    await this.createTarget();
    await this.getMessages();
    await this.getImages();
    await this.createHtml();
  }

  createTarget() {
    return new Promise((resolve) => {
      function probe(location, callback) {
        fs.access(location, err => {
          if (err) {
            probe(path.dirname(location), err => {
              if (err) return callback(err);
              fs.mkdir(location, callback);
            });
          } else {
            callback();
          }
        });
      }
      probe(path.resolve(this.target), resolve);
    });
  }

  readState() {
    try {
      return JSON.parse(fs.readFileSync(path.resolve(this.target, '_state.json'), 'utf8'));
    } catch {
      return null;
    }
  }

  writeState(state) {
    fs.writeFileSync(
      path.resolve(this.target, '_state.json'),
      JSON.stringify(state, null, 2),
      'utf8'
    );
  }

  loadKnownIds() {
    const ids = new Set();
    try {
      const files = fs.readdirSync(this.target).filter(f => FILENAME_MATCH.test(f));
      for (const f of files) {
        const data = JSON.parse(fs.readFileSync(path.resolve(this.target, f), 'utf8'));
        for (const m of data) ids.add(m.id);
      }
    } catch {}
    return ids;
  }

  async getMessages() {
    const state = this.readState();
    const knownIds = this.loadKnownIds();
    const isIncremental = knownIds.size > 0;

    const filePrefix = isIncremental
      ? `delta-${new Date().toISOString().replace(/[:.]/g, '-')}`
      : 'messages';

    console.log(`  ${isIncremental ? `Incremental sync (${knownIds.size} msgs known)` : 'Full backfill'}`);

    let url = `https://graph.microsoft.com/beta/me/chats/${this.chatId}/messages`;
    let page = 0;
    let newestSeen = state ? state.lastSyncedAt : null;
    let totalNew = 0;
    let stopFlag = false;

    while (url && !stopFlag) {
      const pageNum = `${page++}`.padStart(5, '0');
      try {
        console.log(`  Page ${pageNum}`);
        const res = await this.instance.get(url);

        if (res.data.value && res.data.value.length) {
          const newMsgs = [];
          for (const m of res.data.value) {
            if (isIncremental && knownIds.has(m.id)) {
              // Reached known territory. But an edit changes lastModifiedDateTime
              // while keeping the id — those we DO want. Keep any msg whose
              // lastModifiedDateTime is newer than what we knew.
              const t = m.lastModifiedDateTime || m.createdDateTime;
              if (state && state.lastSyncedAt && t > state.lastSyncedAt) {
                newMsgs.push(m);
                if (!newestSeen || t > newestSeen) newestSeen = t;
                continue;
              }
              // Truly old and unchanged — stop paginating.
              stopFlag = true;
              break;
            }
            newMsgs.push(m);
            const t = m.lastModifiedDateTime || m.createdDateTime;
            if (!newestSeen || t > newestSeen) newestSeen = t;
          }

          if (newMsgs.length) {
            await fsAPI.writeFile(
              path.resolve(this.target, `${filePrefix}-${pageNum}.json`),
              JSON.stringify(newMsgs, null, '  '),
              'utf8'
            );
            totalNew += newMsgs.length;
          }
        }

        if (!stopFlag) {
          url = res.data['@odata.nextLink'] || null;
        }
      } catch (err) {
        const status = err.response && err.response.status;
        if (status === 401) {
          console.log('  Hit 401, refresh the token');
          throw err;
        } else if (status === 429) {
          const retry = parseInt((err.response.headers || {})['retry-after'], 10) || 10;
          console.log(`  Hit 429, waiting ${retry}s`);
          await new Promise(r => setTimeout(r, retry * 1000));
          page--;
        } else {
          console.log('  Unhandled error:', err.message);
          break;
        }
      }
    }

    if (newestSeen) {
      this.writeState({
        lastSyncedAt: newestSeen,
        updatedAt: new Date().toISOString()
      });
    }

    console.log(`  ${totalNew} new msgs`);
  }

  async getPages() {
    const filenames = await fsAPI.readdir(this.target);
    return filenames.filter(f => FILENAME_MATCH.test(f)).sort();
  }

  async getImages() {
    const pages = await this.getPages();
    let index = {};
    try {
      const existing = await fsAPI.readFile(path.resolve(this.target, 'images.json'), 'utf8');
      index = JSON.parse(existing);
    } catch {}

    let imageIdx = Object.keys(index).length;

    for (const page of pages) {
      const data = await fsAPI.readFile(path.resolve(this.target, page), 'utf8');
      const messages = JSON.parse(data);

      for (const message of messages) {
        if (!message.body || message.body.contentType !== 'html') continue;
        const imageUrls = message.body.content.match(UPLOADED_IMAGE_MATCH);
        if (!imageUrls) continue;

        for (const imageUrl of imageUrls) {
          if (index[imageUrl]) continue;
          const targetFilename = 'image-' + `0000${imageIdx++}`.slice(-5);
          const imagePath = path.resolve(this.target, targetFilename);

          if (fs.existsSync(imagePath)) {
            index[imageUrl] = targetFilename;
            continue;
          }

          console.log('  Image', targetFilename);
          try {
            const res = await this.instance({ method: 'get', url: imageUrl, responseType: 'stream' });
            res.data.pipe(fs.createWriteStream(imagePath));
            await pipeDone(res.data);
            await new Promise(r => setTimeout(r, 500));
            index[imageUrl] = targetFilename;
          } catch (err) {
            const status = err.response && err.response.status;
            if (status === 403) {
              console.log('  Image 403 (gone)');
              index[imageUrl] = targetFilename;
            } else if (status === 429) {
              const retry = parseInt((err.response.headers || {})['retry-after'], 10) || 10;
              console.log(`  Image 429, wait ${retry}s`);
              await new Promise(r => setTimeout(r, retry * 1000));
            } else {
              console.log('  Image error:', err.message);
            }
          }
        }
      }
    }

    await fsAPI.writeFile(path.resolve(this.target, 'images.json'), JSON.stringify(index), 'utf8');
  }

  async createHtml() {
    const profile = await this.instance.get('https://graph.microsoft.com/v1.0/me/');
    const myId = profile.data.id;

    const pages = await this.getPages();
    let imageIndex = {};
    try {
      imageIndex = JSON.parse(await fsAPI.readFile(path.resolve(this.target, 'images.json'), 'utf8'));
    } catch {}

    // Collect all messages, dedupe by id, keep the most recently modified version.
    const byId = new Map();
    for (const page of pages) {
      const data = await fsAPI.readFile(path.resolve(this.target, page), 'utf8');
      const messages = JSON.parse(data);
      for (const m of messages) {
        const existing = byId.get(m.id);
        const tNew = m.lastModifiedDateTime || m.createdDateTime;
        const tOld = existing && (existing.lastModifiedDateTime || existing.createdDateTime);
        if (!existing || tNew > tOld) byId.set(m.id, m);
      }
    }

    // Sort chronologically (oldest first).
    const all = Array.from(byId.values()).sort((a, b) => {
      const ta = a.createdDateTime || '';
      const tb = b.createdDateTime || '';
      return ta < tb ? -1 : ta > tb ? 1 : 0;
    });

    const fd = await fsAPI.open(path.resolve(this.target, 'index.html'), 'w');
    await fsAPI.write(fd, `<html>
  <head>
    <link rel="stylesheet" href="../../messages.css">
    <meta charset="utf-8">
  </head>
  <body>
`);

    for (const message of all) {
      if (!message.from) continue;
      const timestamp = message.lastModifiedDateTime || message.createdDateTime;

      if (message.from.user != null) {
        const side = message.from.user.id === myId ? 'message-right' : 'message-left';
        await fsAPI.write(fd, `<div class="message ${side}">
  <div class="message-timestamp">${timestamp}</div>
  <div class="message-sender">${escapeHtml(message.from.user.displayName || '')}</div>
`);
        if (message.body.contentType === 'html') {
          await fsAPI.write(fd, `<div class="message-body">${replaceImages(message.body.content, imageIndex)}</div>
</div>
`);
        } else {
          await fsAPI.write(fd, `<div class="message-body">${escapeHtml(message.body.content || '')}</div>
</div>
`);
        }
      } else if (message.from.application != null) {
        await fsAPI.write(fd, `<div class="message message-left">
  <div class="message-timestamp">${timestamp}</div>
  <div class="message-sender">${escapeHtml(message.from.application.displayName || '')}</div>
</div>
`);
      }
    }

    await fsAPI.write(fd, `</body>
</html>
`);
    await fsAPI.close(fd);
    console.log(`  HTML: ${all.length} msgs total`);
  }
}

function escapeHtml(unsafe) {
  return String(unsafe || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function replaceImages(content, imageIndex) {
  if (!imageIndex) return content;
  return content.replace(UPLOADED_IMAGE_MATCH, url => imageIndex[url] || url);
}

function pipeDone(readable) {
  return new Promise((resolve) => { readable.on('end', resolve); });
}

module.exports = Backup;
