# Setup The Keys

In order to setup the form keys find the link to your sheets. The link should look something like this

`https://docs.google.com/spreadsheets/d/18sdlfZvdXlFgYVoHpcwtVUZ6lfdstG7r5KKs5KlmgACfN8/edit?resourcekey#gid=1276308311`

Now that you have your link extract the spreadsheet number. For instance with the link above the spreadsheet number would be `18sdlfZvdXlFgYVoHpcwtVUZ6lfdstG7r5KKs5KlmgACfN8`. The FormKey.txt should contain your form excel sheet (ensure shared link allows viewing). The SheetToFill.txt contains the sheet that you would like to be updated. 

# Client Secret

The client secret can be optained by navigating to the google cloud page, and setting up a new service key. From this a file should be downloaded, if not already done, rename this file to client_secret.json

