# teams-chat-sync

Incremental backup of Microsoft Teams chats via Microsoft Graph API. Lists every chat the authenticated user has access to, downloads all messages and uploaded images, and renders each chat as a browsable HTML file. On subsequent runs it only fetches new/edited messages.

Based on [btecu/teams-chat-backup](https://github.com/btecu/teams-chat-backup), extended with:

- Automatic discovery of all chats (no manual `chatId` list).
- Stable `chatId → folder` mapping so folders survive across runs.
- Incremental sync: after the first backfill, later runs only download new or edited messages.
- Deduplication by message id when regenerating the HTML (edits keep the latest version).
- Fixed pagination stop condition (previously depended on `@odata.count === 20`).
- `Retry-After` honored on 429s.
- Per-chat progress and summary at the end.

## Requirements

Node.js 18+. Delegated `Chat.Read` permission on your Microsoft 365 tenant (see below).

## Install

```sh
git clone https://github.com/Joancesar/teams-chat-sync
cd teams-chat-sync
npm install
```

## Get an access token

1. Open [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) and sign in with your work account.
2. Open the **Modify Permissions** tab, find `Chat.Read`, and consent. If your tenant requires admin approval you will not be able to proceed on your own — ask your admin.
3. Once signed in with the permission granted, open the **Access token** tab and copy the JWT.

Tokens expire in ~60 minutes. If a run dies with a 401, grab a fresh token and rerun — already-synced chats are skipped past quickly and unfinished ones resume where they left off.

## Run

```sh
npm run start
```

Paste the JWT when prompted. The script will:

1. Fetch the list of all chats via `GET /me/chats`.
2. Assign each new chat a folder name (from `topic` for groups, from the other participant for 1:1s, or a chatId suffix as fallback), stored in `out/_chats.json`.
3. For each chat, do a full backfill on the first run or an incremental sync on later runs.
4. Download any new inline images.
5. Regenerate `index.html` from all page files (backfill + deltas).

## Output

```
out/
  _chats.json                   Global chatId → folder mapping
  <chat-folder>/
    _state.json                 { lastSyncedAt, updatedAt }
    messages-00000.json         Backfill pages (page 0 = most recent)
    messages-00001.json
    ...
    delta-<timestamp>-00000.json  Incremental pages from later runs
    ...
    image-00000                 Downloaded inline images
    images.json                 Image URL → local filename mapping
    index.html                  Rendered chat (open in a browser)
```

Open `index.html` in any browser — it links to `messages.css` at the repo root for styling.

## How the incremental sync works

On the first run of a chat there is no `_state.json`, so the script paginates the whole history from newest to oldest and writes each page to `messages-NNNNN.json`. It records the newest `lastModifiedDateTime` seen in `_state.json`.

On later runs the script loads every id from the existing `messages-*.json` and `delta-*.json` files into a `Set`. It paginates from newest again, but as soon as it finds an id already in the Set (and not modified since the last sync) it stops — everything before that point is already local. New messages and edits go into a fresh `delta-<timestamp>-NNNNN.json`.

When the HTML is regenerated, all page files are merged and deduplicated by id. If a message appears more than once (because it was edited), the version with the latest `lastModifiedDateTime` wins.

Deleted messages are not removed from the local backup — a backup is a copy, not a mirror.

## Caveats

- **Compliance**: bulk extraction of corporate chats via Graph is logged in your tenant and may be flagged by DLP policies. Confirm your organization's rules before running this against a work account.
- **Rate limits**: on very large histories you may hit 429s. The script honors `Retry-After`; be patient.
- **Token expiry**: the script does not refresh tokens. On 401 it stops cleanly and you rerun.
- **Beta endpoint**: message fetching uses `/beta/me/chats/{id}/messages`. Microsoft may change its shape without notice.

## License

MIT. Original work © Edgar de Graaff (btecu/teams-chat-backup), extensions © Joan César García León.
