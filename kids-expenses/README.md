# Kids' Shared Expenses

A simple page for the boys' costs that are split 50/50. Alexandra enters the full amount she paid, the page works out John's half, keeps a running total of what he still owes, and keeps a photo or PDF of each receipt that John can open from the item.

## Where the data lives

- Everything is saved in a new **Expense Log** tab in the **KIDS 2026** Google Sheet (the one John can already view). The older tabs are only read, never changed.
- The "JOHN OWES ALEX" total on the **Owed Expenses (Split Amt)** tab is shown as its own line, "Earlier balance", and added to the total. To leave it out, add the Script Property `INCLUDE_OLD_BALANCE` set to `no`.
- The Expenses tab groups everything into Medical, School, Activities and Other dropdowns, like the older tab. Each one lists the items the older tab still counts as owed (the cells its "JOHN OWES ALEX" formula adds up) plus anything new added here.
- The **Coming up** tab shows the next 6 weeks of the boys' appointments, school days off and activities. It reads Alexandra's Google calendar (the one the Skylight uses) and the Family Home Base calendar. From her own calendar it only takes events that name one of the boys, look medical or school related, or that John is invited to. Work events (weddings, clients, video calls) and nanny shifts are always left out.
- Receipts go to a private **Kids Expense Receipts** folder in Alexandra's Drive. They're shown only through the page, so nobody gets a link to the files themselves.
- This repository holds only the code: no names, amounts or receipts.

## Homepage

The page opens on a homepage like the other apps: hello, the date, the balance, then menu squares for Expenses, Add an expense, Coming up and each category.

Alexandra's view also shows when John last opened the page ("John last looked …"). The script notes the time whenever his passcode loads the list and keeps it in Script Properties (`JOHN_LAST_SEEN`), not in the shared Sheet. A Refresh bar along the bottom pulls the latest on every tab.

## Two passcodes

- **Alexandra:** add, edit, mark paid or "John said no", delete.
- **John:** view the list, the totals and the receipts, and add the expenses he paid himself (with a receipt). He can edit or delete only his own items. Alexandra owes him half of those, and the total at the top nets the two sides.

The passcodes are kept in the script's settings and written to a separate private Sheet called **Kids Expenses Passcodes**, never to KIDS 2026 (John can see that one). After 20 wrong tries the page locks for 15 minutes.

## One-time Google setup

1. Open **KIDS 2026** and choose **Extensions > Apps Script**.
2. Paste in everything from `apps-script/Code.gs` and click Save.
3. Pick **setup** in the function menu and click **Run**. Approve Google's permission prompt (if it says "Google hasn't verified this app," click **Advanced**, then **Go to project**).
4. Click **Deploy > New deployment**, type **Web app**, **Execute as: Me**, **Who has access: Anyone**, then **Deploy**.
5. Send the **Web app URL** to Claude, who puts it in `index.html` as `API_URL`.

## Updating the backend later

If a new version reads something new (the calendar, for example), run any function once in the editor first, such as **setup**, so Google can ask for the extra permission.


Paste the new code, then **Deploy > Manage deployments**, edit the deployment and pick **New version**, so the URL stays the same.
