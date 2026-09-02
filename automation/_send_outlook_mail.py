# ── _send_outlook_mail.py ── Runs Outlook COM send in its own process ──────
# Invoked by run.py's send_email() via subprocess with a timeout, so a stuck
# Outlook COM call (e.g. a dialog waiting for input) can be killed instead of
# hanging the whole weekly automation.
import sys, json

import win32com.client

# Outlook shows an inline image only if the attachment carries a Content-ID
# that the HTML then references as cid:<value>. That ID lives in a MAPI
# property with no wrapper in the COM API, hence the raw proptag.
_PR_ATTACH_CONTENT_ID = 'http://schemas.microsoft.com/mapi/proptag/0x3712001F'
# Without this the logo also shows up as a second paperclip attachment
# alongside the workbook.
_PR_ATTACHMENT_HIDDEN = 'http://schemas.microsoft.com/mapi/proptag/0x7FFE000B'


def main(payload_path):
    with open(payload_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = "; ".join(data['to'])
    mail.Subject = data['subject']
    mail.HTMLBody = data['html_body']

    inline = data.get('inline_image')
    if inline:
        try:
            att = mail.Attachments.Add(inline['path'])
            att.PropertyAccessor.SetProperty(_PR_ATTACH_CONTENT_ID,
                                             inline['cid'])
            att.PropertyAccessor.SetProperty(_PR_ATTACHMENT_HIDDEN, True)
        except Exception as e:
            # A footer logo is never worth failing the send over.
            print(f"inline image skipped: {e}", file=sys.stderr)

    if data.get('attachment'):
        mail.Attachments.Add(data['attachment'])
    mail.Send()


if __name__ == '__main__':
    main(sys.argv[1])
