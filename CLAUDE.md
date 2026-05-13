# JOIN Απογραφή — Project Guide

## What this is
A mobile-first web app for monthly inventory counting across 5 JOIN stores. Staff count items, enter quantities on their phones, and export results back to Excel. Devices at the same store sync in real-time; stores are fully isolated from each other.

## Files
- `index.html` — app shell, loads Supabase JS v2 and SheetJS via CDN
- `app.js` — all logic (~1400 lines); 192 items embedded at the top
- `styles.css` — mobile-first CSS
- `supabase-setup.sql` — run once in Supabase SQL Editor to create all tables

## Stores
```javascript
const STORE_NAMES = {1:'Cosmos', 2:'Πατρών', 3:'Λευκός', 4:'Ποσειδώνιο', 5:'OneSalica'};
```

## Supabase tables
| Table | Purpose |
|---|---|
| `inventory_counts` | One row per store/date/item-slot; quantity stored as NET kg |
| `container_weights` | Tare (container) weight per store per ΚΙΛΟ item |
| `item_config` | num_inputs + tare_count per store per ΚΙΛΟ item |
| `store_codes` | SHA-256 hashed store access codes (key: `store_1`..`store_5`, `master`) |

## Key constants & state (app.js)
```javascript
const CATEGORIES = ['FOOD','COFFEE','JUICES','ΑΝΑΛΩΣΙΜΑ','ΤΣΑΪ','ΣΥΣΚΕΥΑΣΙΕΣ'];
const CAT_LABELS = { 'FOOD':'🥗 FOOD', ... };   // display labels with emoji
const ITEM_MAP   = {};   // item_row → item object, built from ITEMS array

let counts          = {};  // { item_row: totalNetQty }
let containerWeights = {}; // { item_row: weight_kg }
let itemConfig      = {};  // { item_row: { num_inputs, tare_count } }
let slotValues      = {};  // { item_row: [netVal0, netVal1, ...] }
let storeCodes      = {};  // { 'store_1': hash, ..., 'master': hash }
```

## Multi-slot system
ΚΙΛΟ items can have multiple input slots (different storage locations). Each slot is saved as a separate row in `inventory_counts` using an encoded `item_row`:
- Slot 0 → stored as `item_row` (original, backward-compatible)
- Slot k (k≥1) → stored as `item_row + k * 10000`

Helper functions: `slotRow(itemRow, slot)`, `parseSlotRow(storedRow)`, `getItemCfg(itemRow)`

`getItemCfg` defaults: `num_inputs=1`, `tare_count=1` if container weight exists else `0`.

## Tare (αποβαρο) logic
- Slots 0..tare_count-1 show a **gross** input ("Μεικτό") + live net display
- `slotValues` always stores **NET** quantities
- Display value for tare slot = `net + cw` (gross)
- When `tare_count` changes, stored values are reinterpreted:
  - plain→tare: `new_net = old_stored - cw`
  - tare→plain: `new_net = old_stored + cw`

## SHA-256 hashing
```javascript
async function sha256(str) {
  // ⚠️ appends 'join-salt' before hashing — always hash with salt
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str + 'join-salt'));
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2,'0')).join('');
}
```
When generating hashes for SQL INSERTs use:
```bash
echo -n "${code}join-salt" | shasum -a 256 | awk '{print $1}'
```

## Access codes
- Store codes: 4-digit PIN per store, verified against `store_codes` table (hashed)
- Master code: works for any store; also unlocks all-store code management in admin panel
- Session auth cached in `sessionStorage` as `join_auth_${n}` = `'1'`
- Auth type cached as `join_auth_type_${n}` = `'store'` or `'master'`
- Default codes (fallback if no DB row): stores `1111`–`5555`, master `0000`

## Admin panel
- Protected by 4-digit admin PIN (default `1234`), stored as SHA-256 hash in `localStorage`
- Shows only for the current store's container weights
- "Κωδικοί Πρόσβασης" section: store entry shows own code only; master entry shows all stores + master
- `saveContainerWeights()` saves `container_weights` table only (item_config is saved per-item on the counting screen via `saveItemCfg()`)

## Export
- Reproduces all 214 rows of the original Excel template exactly
- Column D = counted quantities; column E = `=C{n}*D{n}` formula strings
- Uses `counts[item.row]` (total net) — no slot-level data in export
- Filename: `JOIN_Απογραφή_{StoreName}_{YYYY-MM-DD}.xlsx`

## Workflow
- After every set of changes, update `next_Commit.md` in the project root with a summary of what was changed and why — this file is used as the basis for GitHub commit messages.
- Keep entries concise: one section per logical change, bullet-point format.

## Things to avoid
- Do NOT change the `CATEGORIES` array values — items filter by `i.category === cat`; use `CAT_LABELS` for display only
- Do NOT change slot encoding offsets (10000) without migrating existing DB data
- The `sha256()` function uses `'join-salt'` — never hash without it or stored codes will break
- `refreshTabBadges()` must NOT call `scrollIntoView` — that caused the page-jump-on-typing bug
