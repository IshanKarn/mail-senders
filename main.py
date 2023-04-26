import os, base64
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition

if __name__=="__main__":
    message = Mail(
        from_email='ishan.karn@sageuniversity.in',
        to_emails='ishankarn14@gmail.com',
        subject='Sending with Twilio SendGrid is Fun',
        html_content='<strong>and easy to do anywhere, even with Python</strong>')

    file = "test_doc.xlsx"
    filename = "test_doc.xlsx"
    with open(file, 'rb') as f:
        data = f.read()
        f.close()
    encoded_file = base64.b64encode(data).decode()
    attached_file = Attachment(
        FileContent(encoded_file),
        FileName(filename),
        FileType('application/xlsx'),
        Disposition('attachment')
    )

    message.attachment = attached_file
    try:
        sg = SendGridAPIClient(os.environ.get('SENDGRID_API_KEY'))
        response = sg.send(message)
        print(response.status_code)
        print(response.body)
        print(response.headers)
    except Exception as e:
        print(e.message)