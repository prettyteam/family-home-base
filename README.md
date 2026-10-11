# Family Home Base

A shared organizer for Alexandra and the family's nanny: nanny shifts and hours, the kids' assignments and grades, weekly activities, and to-dos. It runs as a plain web page (GitHub Pages), so nobody needs a Claude account to use it.

## How it keeps family details private

- This repository holds only the app's code. No names, schedules, grades or pay are stored here.
- All family data lives in a private Google Sheet owned by Alexandra.
- The page can only read or change that data with a passcode. There are two passcodes:
  - **Owner passcode:** full access.
  - **Nanny passcode:** she can see the schedule, clock in and out, fix her own times, and check off to-dos and homework. She can't change pay, grades, activities or house notes, and she only sees pay if that's turned on.
- After 20 wrong passcode attempts, the app locks for 15 minutes.
- Never put school or account passwords in the app.

## One-time Google setup (about 5 minutes)

The Google Sheet **Family Home Base Data** already exists in Alexandra's Drive.

1. Open the Sheet and choose **Extensions > Apps Script**.
2. Delete what's there, paste in everything from `apps-script/Code.gs`, and click Save.
3. In the function menu at the top, pick **setup** and click **Run**. Approve Google's permission prompt. (If it says "Google hasn't verified this app," click **Advanced**, then **Go to project**.) This creates both passcodes in a **Passcodes** tab of the Sheet and turns on the Skylight calendar.
4. Click **Deploy > New deployment**. For type, choose **Web app**. Set **Execute as: Me** and **Who has access: Anyone**, then click **Deploy**.
5. Copy the **Web app URL** and send it to Claude.

"Anyone" lets the page reach the script without a Google login. The passcode is what protects the data.

## Showing it on the Skylight

The app can copy the kids' weekly activities and the nanny's shifts for the next 4 weeks into a Google Calendar named **Family Home Base**. The Skylight can then show that calendar.

1. The `setup` step above turns this on. It creates the calendar and refreshes it every 15 minutes when something changes, and once a day.
2. In the Skylight app, open calendar sync, connect the same Google account, and turn on the **Family Home Base** calendar.
3. To keep nanny shifts off the Skylight, add the Script Property `SKYLIGHT_SHOW_SHIFTS` set to `no`.

The sync only goes one way, from the app to the Skylight. Events added on the Skylight don't come back into the app.

## Updating the backend later

After you paste new code into Apps Script, choose **Deploy > Manage deployments**, edit the existing deployment, and pick **New version**. Doing it this way keeps the same URL.
