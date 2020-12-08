import base64

from apiclient import errors
from email.mime.text import MIMEText
from google.oauth2 import service_account
from googleapiclient.discovery import build

EMAIL_FROM = 'eanderson@khitconsulting.com'


def create_message(sender, to, subject, message_text):
    """ Create a message for an email.

        Args:
            sender: Email address of the sender.
            to: Email address of the receiver.
            subject: The subject of the email message.
            message_text: The text of the email message.

        Returns:
            An object containing a base64url encoded email object.
    """
    message = MIMEText(message_text)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    return {'raw': base64.urlsafe_b64encode(message.as_string().encode('utf-8'))}


def send_message(service, user_id, message):
    """ Send an email message.

        Args:
            service: Authorized Gmail API service instance.
            user_id: User's email address. The special value "me"
            can be used to indicate the authenticated user.
            message: Message to be sent.

        Returns:
            Sent Message.
    """
    try:
        message['raw'] = message['raw'].decode()
        message = (service.users().messages().send(userId=user_id, body=message).execute())
        print('Message Id: %s' % message['id'])
        return message
    except errors.HttpError as error:
        print('An error occurred: %s' % error)


def service_act_login():
    """ Log into gmail service account. """
    SCOPES = ['https://www.googleapis.com/auth/gmail.send']
    SERVICE_ACCOUNT_FILE = 'src/config/service_key.json'

    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    delegated_credentials = credentials.with_subject(EMAIL_FROM)
    service = build('gmail', 'v1', credentials=delegated_credentials)
    return service


def send_gmail(email_to, subject, content):
    service = service_act_login()
    signature = '\n\n--\nKHIT Consulting\nwww.khitconsulting.com\n\nThis email and its attachments may contain ' \
                'privileged and confidential information and/or protected health information (PHI) intended solely ' \
                'for the use of KHIT Consulting LLC and the recipient(s) named above.  If you are not the recipient, ' \
                'or the employee or agent responsible for delivering this message to the intended recipient, you ' \
                'are hereby notified that any review, dissemination, distribution, printing or copying of this ' \
                'email message and/or any attachments is strictly prohibited.  If you have received this ' \
                'transmission in error, please notify the sender immediately and permanently delete this email ' \
                'and any attachments.'
    content = content + signature
    message = create_message(EMAIL_FROM, email_to, subject, content)
    sent = send_message(service, 'me', message)
    return sent
