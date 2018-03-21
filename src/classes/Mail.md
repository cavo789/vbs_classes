# Mail.vbs

Provide a simple solution to work with emails

## Table of content

- [Send](#send)

## Display

Display an Outlook email message with To, Subject and Body already initialized.

The user will see the message, can modify it if needed and he'll decide to send himself the email (no automatic sending, it's a "display" action)

### Sample script

```vbnet
Set cMail = New clsMail
cMail.Recipient = "test@yahoo.com"
cMail.Subject = "A subject"
cMail.HTMLBody = "Do you ? Really ???"
cMail.Display()
Set cMail = Nothing
```
