# -*- coding: utf-8 -*-
import os
from datetime import datetime
from functools import partial
from pathlib import Path
from typing import List, Union, Tuple
import win32com.client

excluded_folders = ['Deleted Items', 'Junk Email', 'Drafts', 'Conversation History', 'Yammer Root', 'Trash', 'Sent']


def _get_mapi_namespace() -> any:
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")
    return mapi


def _target_acc_exists(target_email: str, mapi) -> bool:
    for account in mapi.Accounts:
        if target_email == account.DeliveryStore.DisplayName:
            return True
    return False


def _iterate_all_inboxes(target_acc, verbose: bool = False) -> any:
    for target_folder in target_acc.Folders:
        if target_folder.Name in excluded_folders:
            continue

        if verbose:
            print(f"Iterating over: {target_folder.Name}")

        yield target_folder.Items


def _get_sender_name(email) -> str:
    try:
        return email.SenderName
    except AttributeError as e:
        pass

    try:
        return email.Organizer
    except AttributeError as e:
        return "Unknown. :("


def _get_sender(email) -> Tuple[Union[str, None], Union[str, None]]:
    sender_name = _get_sender_name(email)

    try:
        sender_email = email.SenderEmailAddress
    except AttributeError as e:
        sender_email = "N/a (Probably a meeting)"

    return sender_name, sender_email


def _get_received_at(email) -> Union[datetime, None]:
    try:
        return email.ReceivedTime
    except AttributeError as e:
        pass

    try:
        return email.CreationTime
    except AttributeError as e:
        pass

    return None


def _sanitize_subject(subject):
    """
    Url: https://gist.github.com/wassname/1393c4a57cfcbf03641dbc31886123b8
    """
    import unicodedata
    import string
    whitelist = "-_.() %s%s" % (string.ascii_letters, string.digits)
    char_limit = 255
    replace = " "

    # replace spaces
    for r in replace:
        subject = subject.replace(r, '_')

    # keep only valid ascii chars
    cleaned_subject = unicodedata.normalize('NFKD', subject).encode('ASCII', 'ignore').decode()

    # keep only whitelisted chars
    cleaned_subject = ''.join(c for c in cleaned_subject if c in whitelist)
    if len(cleaned_subject) > char_limit:
        print("Warning, subject truncated because it was over {}. Filenames may no longer be unique".format(char_limit))
    return cleaned_subject[:char_limit]


def main(target_email: str, attachment_base_folder: Path, email_filter_by_subject: callable, verbose: bool) -> None:
    mapi = _get_mapi_namespace()
    if not _target_acc_exists(target_email, mapi):
        print(f"ERROR: No accounts found with email '{target_email}'")
        return

    if not attachment_base_folder.exists():
        os.makedirs(attachment_base_folder)

    target_acc = mapi.Folders[target_email]
    count = 0
    matched = 0
    for emails in _iterate_all_inboxes(target_acc, verbose):
        for email in emails:
            count += 1
            subject = email.Subject
            if email_filter_by_subject is not None and not email_filter_by_subject(subject):
                continue

            matched += 1
            received_at = _get_received_at(email)
            message = email.Body
            sender, sender_email = _get_sender(email)
            sanitized_subject = _sanitize_subject(subject)
            attachment_folder = attachment_base_folder.joinpath(
                f"{received_at.strftime('%Y.%m.%d')} - {sanitized_subject}")
            if len(email.Attachments) > 0 and not attachment_folder.exists():
                os.makedirs(attachment_folder)
                with open(attachment_folder.joinpath(f"{sanitized_subject}.txt"), "w", encoding="utf-8") as file:
                    file.write(f"From: {sender} <{sender_email}>\n"
                               f"Subject: {subject}\n"
                               f"Received at: {received_at.isoformat()}\n"
                               f"{message}")

            for attachment in email.Attachments:
                if verbose:
                    print(f'\t{attachment.FileName}')
                attachment_filename = attachment_folder.joinpath(attachment.FileName)
                attachment.SaveASFile(str(attachment_filename.absolute()))
            if verbose:
                print(f"[{received_at}] From: {sender}: {subject}")

    print(count, matched)


def email_filter_by_subject_callback(subject: str, target_keywords: List[List[str]]) -> bool:
    _subject = subject.lower()
    return any(all(w.lower() in _subject for w in ts) for ts in target_keywords)


if __name__ == '__main__':
    main(
        target_email="change@me.org",
        attachment_base_folder=Path("attachments"),
        email_filter_by_subject=partial(email_filter_by_subject_callback,
                                        target_keywords=[
                                            ["recibo", "pagamento"],
                                            ["informe", "rendimento"],
                                            ["imposto", "renda"],
                                            ["irpf", ]
                                        ]),
        verbose=True
    )
