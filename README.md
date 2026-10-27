# Inventory Check Out

A check-out/check-in page for the **Feehan-DM Inventory** Google Sheet. The page is hosted on GitHub Pages, and a Google Apps Script handles the sheet.

- Check-outs and check-ins are written to **Form Responses 1** in the Google Form's own format (`Take out` / `Check in`), so each tab's **Total Out** and **Current Qty** formulas count them along with real form submissions.
- Items and their Current Qty are read live from the OFSA, DM Office and Warehouse tabs. Warehouse is link only: it has no tab button, and opens when `?office=Warehouse` is added to the site link. The item columns each tab uses in Form Responses 1 are read from its Total Out formula.
- "Checked out now" lists reusable items people still have (take-outs minus check-ins). Consumables are counted but never listed as checked out.
- For speed, one request loads every tab, the server keeps that result for 20 seconds (any write clears it), and the page shows its last copy instantly while it refreshes.

## Admin portal

The **Admin** link at the bottom of the site opens the admin portal. You can change everything there without editing any code:

- **Checkouts:** see who has what, check items in for someone, and search or delete rows in Form Responses 1.
- **Inventory:** edit, add or delete item rows (the Total Out / Current Qty formulas are left alone, and new rows get them).
- **Names:** NetIDs show as real names, looked up in the Rutgers Google directory (People API) or taken from what people typed; fix any name under **Settings → People**.
- **Settings:** set the site title, choose whether an event is required, and manage offices (add, rename, reorder, point them at a different sheet tab and starting row, or make one link only).

**Who can open the portal:** the people on the admin list in **Settings → Admins**. Admins add and remove each other there. If someone you add doesn't have edit access to the sheet yet, adding them gives it to them, because the portal needs it. If the list is ever empty, every editor of the sheet can get in.

## Apps Script setup

The Apps Script project has two files, [`apps-script/Code.gs`](apps-script/Code.gs) and [`apps-script/Admin.html`](apps-script/Admin.html). It is deployed twice:

| Deployment | Execute as | Who has access | Used by |
|---|---|---|---|
| Public API | Me | Anyone | `API_URL` in `config.js` |
| Admin portal | User accessing the web app | Anyone with Google account | `ADMIN_URL` in `config.js` (with `?page=admin`) |

After you change the script, update **both** deployments with **Deploy → Manage deployments → Edit → Version: New version**. Their URLs stay the same.
