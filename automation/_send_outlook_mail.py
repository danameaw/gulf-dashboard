# ── _send_outlook_mail.py ── Runs Outlook COM send in its own process ──────
# Invoked by run.py's send_email() via subprocess with a timeout, so a stuck
# Outlook COM call (e.g. a dialog waiting for input) can be killed instead of
# hanging the whole weekly automation.
import sys, json

import win32com.client


def main(payload_path):
    with open(payload_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = "; ".join(data['to'])
    mail.Subject = data['subject']
    mail.HTMLBody = data['html_body']
    if data.get('attachment'):
        mail.Attachments.Add(data['attachment'])
    mail.Send()


if __name__ == '__main__':
    main(sys.argv[1])
