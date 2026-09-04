const fs = require('fs');
const path = require('path');
const readline = require('readline');
const axios = require('axios');
const Backup = require('./backup');

const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

function ask(question) {
  return new Promise((resolve, reject) => {
    rl.question(`${question} `, answer => {
      if (!answer.trim()) return reject(new Error('missing value'));
      resolve(answer);
    });
  });
}

function sanitize(s) {
  return (s || '').replace(/[^\w\-]+/g, '_').slice(0, 60) || 'chat';
}

const CHATS_MAP = path.resolve('out/_chats.json');

function loadMap() {
  try {
    return JSON.parse(fs.readFileSync(CHATS_MAP, 'utf8'));
  } catch {
    return {};
  }
}

function saveMap(map) {
  fs.mkdirSync(path.dirname(CHATS_MAP), { recursive: true });
  fs.writeFileSync(CHATS_MAP, JSON.stringify(map, null, 2), 'utf8');
}

async function listAllChats(authToken) {
  const headers = { Authorization: `Bearer ${authToken}` };
  let url = 'https://graph.microsoft.com/v1.0/me/chats?$top=50&$expand=members';
  const all = [];
  process.stdout.write('Listing chats');
  while (url) {
    const res = await axios.get(url, { headers });
    all.push(...res.data.value);
    url = res.data['@odata.nextLink'];
    process.stdout.write('.');
  }
  process.stdout.write(`\n${all.length} chats found\n`);

  const map = loadMap();
  const conversations = [];

  for (const c of all) {
    if (!map[c.id]) {
      let name;
      if (c.topic) {
        name = c.topic;
      } else if (c.chatType === 'oneOnOne' && c.members) {
        const other = c.members.find(m => m.displayName);
        name = other ? other.displayName : `oneonone-${c.id.slice(-12)}`;
      } else {
        name = `${c.chatType || 'chat'}-${c.id.slice(-8)}`;
      }
      // Deduplicate target names
      let target = sanitize(name);
      const used = new Set(Object.values(map));
      let suffix = 2;
      const base = target;
      while (used.has(target)) {
        target = `${base}_${suffix++}`;
      }
      map[c.id] = target;
    }
    conversations.push({ chatId: c.id, target: map[c.id] });
  }

  saveMap(map);
  return conversations;
}

async function main() {
  const authToken = await ask('Enter JWT:');
  const conversations = await listAllChats(authToken);

  let ok = 0;
  let failed = 0;

  for (let i = 0; i < conversations.length; i++) {
    const c = conversations[i];
    console.log(`\n[${i + 1}/${conversations.length}] ${c.target}`);
    const backup = new Backup({
      chatId: c.chatId,
      authToken,
      target: `out/${c.target}`
    });
    try {
      await backup.run();
      ok++;
    } catch (err) {
      failed++;
      console.error(`  ERROR: ${err.message}`);
      if (err.response && err.response.status === 401) {
        console.error('  Token expired. Stopping. Re-run with a fresh token.');
        break;
      }
    }
  }

  console.log(`\nSummary: ${ok} ok, ${failed} failed, out of ${conversations.length}`);
}

main()
  .then(() => rl.close())
  .catch(err => {
    rl.close();
    console.error(err);
    process.exit(1);
  });
