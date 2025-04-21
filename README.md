# GmailTaxAutomation
A Google App Script automation which will move any emails and attachments labelled with a "tax" label, into a Google Drive folder called "Tax Deductions".
From there, Microsoft Power Automate can pick up the files and move them into a OneDrive file.

I developed this as there was no simple way to trigger a Microsoft Power Automate flow from a gmail label being applied.


## Clasp
Instructions to set up CLASP so you can push the code to Google App Script.
1. install CLASP eg "brew install clasp" on mac
2. clasp login
3. clasp push <script id>
