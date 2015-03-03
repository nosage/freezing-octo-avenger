# freezing-octo-avenger
This will run a report of mailbox sizes, mailbox quotas, database quotas and percentage of the mailbox size that is being used.
It will first save the report as an .html to c:\scripts and then email the report in the body of the email using the email settings you need to provide.

This script assumes you have a c:\scripts directory to save the reports.
This script assumes the mailboxes have been initialized (logged on to).
Be aware how many days of html reports piles up in your c:\scripts folders, purge as needed to prevent C: from filling up.

There are a lot of get-mailbox and get-mailboxdatabase requests which makes this a very demanding script, I wish I can lighten it up but there is a ton of data I want in my report.  With that in mind I would not recommend this for mailbox servers that have thousands of mailboxes.  If you admin one of those you probably have a better reporting tool in place anyway.

Tested using Exchange 2013, should work on 2010, not sure, sorry!
