# JS plugin for Outlook - get attachment details and modify email body

## About

This project is a sample JS plugin for Outlook created to illustrate the basic aspects of creating ones. It receives information about attachments when user tries to send an email and adds these details to the body of the email.

## Implementation

The plugin performs such actions:


- Adds the “Analyze And Send” button that calls custom code before sending an email
- Gets attachment details (AttachmentId, Name, ContentType, Size, LastModifiedTime)
- Adds received data to the email body (modifies email body)
- Adds custom properties to email, which can be then used by receiver
- Initiate sending of the modified email to specific addressees (not only those specified by the sender)

For detailed implementation notes as well as some basic recommendations and COM technology comparison, please go to the [related article](https://www.apriorit.com/dev-blog/431-how-to-develop-javascript-outlook-plugin).

## License

Licensed under the MIT license. © Apriorit.
