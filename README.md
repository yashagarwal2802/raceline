# RaceLine — Order & Stock Desk

A small app for your team to take orders, track billing/dispatch, manage
SKF stock, and track pending orders with SKF. It runs on one office PC and
everyone else connects to it over your WiFi — no cloud, no ongoing cost,
no logins beyond picking your name.

## One-time setup (on the PC that will run it)

1. **Install Node.js** — go to https://nodejs.org, download the "LTS"
   version for your computer, and install it like any other program.
   (This is a one-time step; skip it if Node is already installed.)

2. **Unzip this folder** somewhere convenient, e.g. your Desktop.

3. **Open a terminal in this folder:**
   - **Mac:** right-click the `raceline` folder → "New Terminal at Folder"
     (or open Terminal and type `cd ` then drag the folder in, then Enter)
   - **Windows:** open the folder in File Explorer, click the address bar,
     type `cmd`, press Enter

4. **Install once:**
   ```
   npm install
   ```

5. **Start it:**
   ```
   npm start
   ```

You'll see something like:

```
  RaceLine is running.

  On this PC:      http://localhost:3000
  On the network:  http://192.168.1.42:3000

  Open the network address on any phone or PC on the same WiFi.
  Leave this window open — closing it stops RaceLine.
```

Open the **network** address (the `192.168.x.x` one) on your phone or PC's
browser, and share that same address with the rest of your team — anyone on
the same office WiFi can use it. Leave the terminal window open; closing it
stops RaceLine for everyone. If you restart the PC, just run `npm start`
again from this folder.

## First login

You'll see one person: **Yash Agarwal (Owner)**. Pick that, then go to the
**Team** tab and add your other 6 staff with their roles (Order taker,
Biller, Dispatch). Once added, they'll see their name on the picker screen
from their own phone/PC.

## Replacing the sample data

The **Stock** and **Customers** tabs are pre-loaded with a few sample SKF
part numbers and dealers (marked "Sample") just so the app isn't empty on
first look. Add your real catalog and dealer list (one at a time, or via
"Bulk import" — paste rows copied from Excel), then click **Clear samples**
in each tab to remove the placeholders.

## What's in here

- **New order** — order-takers pick a customer and items, submit the order
- **Orders** — status pipeline: Taken → Billed → Dispatched, with actions
  gated by role; stock reduces only when an order is actually dispatched
- **Stock** — live catalog with location, reorder-level alerts, search
- **Purchases** — log what you've asked SKF for and what's arrived; tracks
  what SKF still owes you per part number (since there's no PO number on
  either side, it's matched by a running per-part ledger, not by order ID);
  logging a receipt also shows which pending customer orders it should
  fulfill first
- **Customers** — dealer list, credit period, and outstanding balance
- **Team** — add/remove staff and set roles
- **Dashboard** (Owner) — today's order counts, low stock, slow movers,
  backup export

## Putting an icon on your phone (instead of typing the link)

Once RaceLine is open in your phone's browser, you can add it to your home
screen like an app — same icon, opens full-screen, no address bar, nothing
to type after that:

- **iPhone (Safari):** tap the Share icon (square with an arrow) → **"Add to
  Home Screen"** → Add
- **Android (Chrome):** tap the ⋮ menu (top right) → **"Add to Home
  screen"** (or you may see an "Install app" prompt directly) → Add/Install

Do this once per phone. Note: this works over your office WiFi's address
today; if you move RaceLine to always-reachable cloud hosting later, the
icon will need to be re-added once pointing at the new address.

## Backing up your data

All data lives in one file: `data/data.json`. Copy it somewhere safe now
and then (a USB drive, email it to yourself, etc.), or use the **Export
backup** link on the Dashboard tab any time.

## Known limits, for now

- Tally is not yet connected — the "Tally invoice number" on a billed order
  is just a note you type in, not verified against Tally itself. Real Tally
  billing figures on the Dashboard, and outstanding-balance auto-sync, are
  planned once we test the connection against your actual Tally setup.
- No photo/handwriting order entry yet — orders are entered manually for now.
