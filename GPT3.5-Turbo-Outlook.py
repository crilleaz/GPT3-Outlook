import win32com.client
import bs4
import openai
import time

openai.api_key = "" # obtained via https://platform.openai.com/account/api-keys

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
signature = "" # your signature, use \n for newline


while True:
    messages = inbox.Items
    found_unread = False
    for message in messages:
        if message.UnRead:
            soup = bs4.BeautifulSoup(message.Body, "html.parser")
            content = soup.get_text().split("From:")[0]

            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                max_tokens=2000,
                n=1,
                stop=None,
                temperature=0.55,
                messages=[
                        {"role": "user", "content": content},
                    ]
            )
            result = ''
            for choice in response.choices:
                result += choice.message.content
            reply = message.Reply()
            reply.Body = result + "\n\n\n" + signature
            reply.Subject = "Sv: " + message.Subject # 'Sv: ' for swedish mails, other countries use 'Re: '
            reply.Send()
            message.UnRead = False
            found_unread = True

    if not found_unread:
        print("Waiting for incoming emails..")

    time.sleep(1200)
