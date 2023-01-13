# email_to_whatsapp
Problem : Currently I am receiving an email containing an attachment from an admin member of an organization every 2 weeks.
I am tasked to forward this attachment to my group mates, and this has previously been done manually via downloading the attachment and then sending on a whatsapp group.
In addition to sending the attachment, I had to send a message to inform them of the details of the meeting i.e location, date, timing as well as to collate attendance.

My initial plan was to write python script to automate the conversion of the email to a whatsapp message, including the attachment, to the group.
Having searched online for the relevant libraries which would allow me to send the file over whatsapp, I realized I needed to convert the (usually docx) file into a jpeg image.
This is usually done using the win32com library but unfortunately being a Mac user I don't have access to that library.

Hence my (current) solution would be to simply automate the downloading of the attachment, and then sending it to my group mates as a seperate email.
The 'usual' whatsapp message is now sent via the pywhatkit library.
