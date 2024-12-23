# OutlookForB2BSales

NUMBER 1
CorrectSpanishLanguageAccentImportedOnOutlookContact
This script allows you to handle non ASCII carachters after importing spanish accented names into outlook.

NUMBER 2
AutomatizzaInvioMailFollowup
This script automates the sending of personalized emails to a list of contacts whose emails have been moved to a specific folder.
The email address to which the follow-up email is sent is identified from the previously sent email, which must be specifically placed in a folder named PROVAMACRO01 within Sent email directory.

NUMBER 3
I often send emails to a list of contacts in Outlook. However, some addresses are incorrect, and I receive emails like this one that end up in a specific folder:
"""The message could not be delivered to example.red@example.com.
acorreia was not found in example.com."""
The initial sentence, "The message could not be delivered to example.red@example.com," contains the email address that is unreachable. 
I created an outlook rule that automatically moves messages with "Non recapitabile" in the object to a BOUNCED folder I created.
This VBA script identifies all the unreachable addresses mentioned in the emails within the BOUNCED folder and matches them with the corresponding contacts in the Contacts folder, flagging or categorizing them in red.
