# GmailTaxAutomation
A Google App Script automation which will move any emails and attachments labelled with a "tax" label, into a Google Drive folder called "Tax Deductions". This will run as a batch job every 24 hours.

I developed this because I wanted to put any gmail mails/attachments with "tax" labelled into Microsoft Onedrive. To work around some limitations in the Microsoft Power Automate connector for gmail,
1) there is way to trigger a Microsoft Power Automate flow from a gmail label being applied
2) there is no "get messages" action. 

By putting the files into Google Drive, I can run a Power Automate flow to pick them up from that folder and move them into a OneDrive folder using the OneDrive connector.

# Setup instructions

## Setup App script

### Get this code
Use git clone to pull this code

### Setup Clasp
You should setup CLASP so you can send this code back/forward to google appscript from the command line
1. install CLASP 
    1. eg `brew install clasp` on mac
2. setup script 
    1. If you already have a script file: `clasp 
clone-script <scriptID>`
    2. If you don't have a script file: `clasp create-script`
3. send the code to the script in appscript `clasp push`

### Set the code to run
Heres how to manually set up the trigger.
1. Go to google app script https://script.google.com/
1. Go to your script 
1. Go to code > choose the createMidnightTrigger and press run. This will create the trigger to run every night.

Not sure how to create it from the command line yet, but I'm sure its possible.