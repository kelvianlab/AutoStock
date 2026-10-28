<div align="center">

# AutoStock

**A warehouse stock app made from one Google Sheet.**
**The admin types what came in and what went out. The stock counts itself.**

No server. No install. No monthly fee. Works on a laptop or a phone.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Google Apps Script](https://img.shields.io/badge/built%20with-Google%20Apps%20Script-4285F4.svg)](https://developers.google.com/apps-script)

</div>

---

## What is this?

Most small warehouses keep stock in a spreadsheet that one person updates by hand.
Everyone else has to ask that person how much is left.

AutoStock turns a normal Google Sheet into a small stock app. The warehouse admin
only records items coming in and going out. Every time a cell changes, the script
recalculates the stock from scratch, so the numbers stay right even after
Ctrl+Z, a deleted row, or a typo that gets fixed later.

Share the file with your team: the admin gets **Editor**, everyone else gets
**Viewer** and sees the live stock without being able to change it.

## The four sheets

| Sheet | What it is for |
|---|---|
| **Master** | Look up items and their current stock |
| **Stock In** | Record incoming items |
| **Stock Out** | Record outgoing items |
| **Archive** | Items whose stock reached 0 |

`setup` creates all four for you.

## Install (once per file, about 5 minutes)

1. In Google Drive, click **New**, then **Google Sheets**, then **Blank spreadsheet**.
2. Open **File**, then **Settings**. Set **Time zone** to yours, then **Save settings**.
3. Open **Extensions**, then **Apps Script**.
4. In `Code.gs`, delete the default `function myFunction()`, paste the whole of
   [`AutoStock.gs`](AutoStock.gs), and save with Ctrl+S.
5. Make sure the function dropdown next to **Debug** shows `setup`, then click **Run**.
6. Allow the permissions. If Google says the app is not verified, that is because
   you wrote it yourself: click **Advanced**, then continue.
7. Wait for **Execution completed**, go back to the sheet and refresh.

A new menu, **🛠 Modification Tools**, appears at the top.

## Daily use

**Existing item comes in:** in **Stock In**, pick the **Item Code** from the
dropdown. Name, spec, type and brand fill in by themselves. Type the **Qty In**.

**New item comes in:** leave **Item Code** empty, fill in Item Name, Size/Spec,
Type, Brand and Unit, then type the **Qty In**. A new code is created for you.
If the name matches an item that already exists (in Master or Archive), its old
code is used instead.

**Item goes out:** in **Stock Out**, pick the **Item Code** and type the **Qty Out**.

The **Status** column turns green `OK` once a row is counted, and red with a
message when something needs fixing.

## What runs by itself

| Feature | Where in the script |
|---|---|
| Recalculate on typing, paste, Ctrl+Z | `onEditHandler` (onEdit trigger) |
| Recalculate when rows are deleted | `onChangeHandler` (onChange trigger) |
| New item codes 0001, 0002, ... | `newCode_` |
| Same item despite case, spaces, dots, commas | `identityKey_` |
| Code dropdown in both logs | `requireValueInRange` in `setup` |
| Red warning when more goes out than is in stock | `rowStatus_` |
| Change a unit anywhere, it changes everywhere | step 7b in `computeStock_` |
| Sold-out items move to Archive and come back with their old code | step 9 in `computeStock_` |
| Menu **Recalculate Stock** | `onOpen` and `recalculateButton` |

## Fixing mistakes

| Mistake | What to do |
|---|---|
| Wrong qty | Type the number again |
| Whole row is wrong | Ctrl+Z, or right-click and **Delete row** |
| Wrong name or spec | Retype it in Master, or in the log row |
| A typo created a duplicate item | Fix the name in the log. It merges, and the extra code disappears |
| Wrong unit | Change it on any sheet. Every sheet follows |

Stock is always rebuilt from the logs, so all of this is safe to repeat.

## Good to know

- **Code** and **Qty** in Master are protected with a warning. If someone types
  there anyway, the script puts the right value back on the next change.
- Not enough stock is **not rejected**, so Ctrl+Z always works. The row turns red
  and Master shows the negative number in red until the qty is fixed.
- Do not rename the four sheets. If you need other names, change them in `CONFIG`
  at the top of the script as well.
- The stock check compares total in with total out. It does not look at the order
  of the dates.
- Free Google accounts have a daily quota for script run time. A normal office
  warehouse stays far below it.
- One file per warehouse location. For two warehouses, make two files.
- If Master ever looks wrong, open **🛠 Modification Tools** and click
  **🔄 Recalculate Stock**. It is safe to click at any time.

## License

[MIT](LICENSE)
