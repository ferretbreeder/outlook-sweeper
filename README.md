# outlook-sweeper
Someone came to me with a request: write something that will take an email from a folder in their Outlook mailbox, append some simple text to the top of it, and save it as a PDF. Simple enough, right? That's the hope! 

## Planned features
There are a few basic requirements for this project:

- run as an unprivileged user as a standalone Python executable that doesn't require any other binaries/programs to run
- run entirely on the local machine, as to remain compliant with various regulations in the industry
- the finished product needs to very closely resemble the result of the manual process that this application is taking the place of

At the end of it, the user will be able to provide a first name, last name, and ID number, and those values will be appended to the top of the PDF. 