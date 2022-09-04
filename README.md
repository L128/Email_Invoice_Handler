# Email_Invoice_Handler
Using IMAP to receive emails, pdfplummer to check the invoices in the attachment, SMTP to send them to my accountant.


Using a lot of code from @Paramiao, https://github.com/paramiao/pyMail.
This pyMail.py is sufficient for basic email scripting.

I added support for Chinese and updated old APIs.


main.py does:
1. Goes over INBOX
2. Collect the invoice(fapiao) using a temperary directory
3. Archive emails with legit invoice (If target folder name accurate in credentials.conf)
4. Plots a summary

**Note. It is important to rename credential-sample.conf to credential.conf. And don't use quotes in .conf.
