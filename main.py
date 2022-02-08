import os
import dotenv
import re
import exchangelib

dotenv.load_dotenv(verbose=True)


def exchange_connect(usr, pwd, smtp_acc):

    # credentials for login
    credentials = exchangelib.Credentials(usr, pwd)

    # set config for MS365 Exchange
    config = exchangelib.Configuration(
        server="outlook.office365.com", credentials=credentials
    )

    # connect to account
    account = exchangelib.Account(
        # smtp_acc could be different to logged in user and will be dependant on permissions
        primary_smtp_address=smtp_acc,
        config=config,
        autodiscover=False,
        access_type=exchangelib.DELEGATE,
    )

    return account


def get_folders(account):
    return account.root.tree()


# this my be different from your user
email_box = "yourmailbox@yourdomain.co.za"
# connect to exchange account
account = exchange_connect(os.getenv("EMLUSER"), os.getenv("EMLPWD"), email_box)

# return folders
print(get_folders(account))

# get list of list 5 e-mails received
for item in account.inbox.all().order_by("-datetime_received")[:5]:
    print(item.subject, item.sender, item.datetime_received)

# super useful use case is for checking undelivarable e-mails
for item in account.inbox.filter(subject__contains="Undeliverable"):
    print("Subject: " + item.subject)
    # filter for e-mails and generate list
    emails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", item.body)
    # remove duplicates from list
    emails = list(dict.fromkeys(emails))
    # remove mailbox e-mail address
    email = emails.remove(email_box)
    print("Failed emails:")
    # list of failed e-mail addresses
    print(emails)
