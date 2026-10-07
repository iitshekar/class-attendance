# Class check-in

Students scan a rotating QR code on the classroom screen with their own tablets. Each tablet is linked to one student, and the teacher approves every new tablet. The teacher reviews absentees, approves, and the day's record is saved to a Google Sheet (Daily Log). Late scans after approval are recorded as late automatically.

No student data is stored in this repo. Names live only in the teacher's Google Sheet.

## One-time setup
1. Create a Google Sheet named "Attendance". Open Extensions > Apps Script, paste `Code.gs`, and save.
2. Run `setup()` and allow access. Copy the **teacher key** from the execution log.
3. In the **Roster** tab, add one row per student: Class | Student (for example 7A | Aisha Khan).
4. In Apps Script, go to Deploy > New deployment > Web app. Set Execute as **Me** and Who has access **Anyone**. Copy the web app link.
5. Open the GitHub Pages link on your laptop or phone and paste in the web app link and teacher key.

## Registration day
Students scan the QR code, choose their name and tap "Link this tablet to me". Approve each request on your screen. Approving also checks them in for that day.

## Daily use
Open the page and show the QR code on the screen. Students scan and see a green tick. Tap "Finish and review absentees", check the list, then tap "Approve and save". Use "Send to my WhatsApp", then switch the listed students to Absent in Nexquare.

## Managing tablets
The **Devices** tab lists every tablet. A new tablet for the same student shows a warning; approving it replaces the old one. To unlink a tablet completely, delete its row.

## Updating the script
After editing `Code.gs`, use Deploy > Manage deployments > Edit > New version so the same web app link keeps working.
