# pytlook: Python-Outlook Attachment Extractor
This script extracts the attachments from emails where the subject matches a certain criteria.
The criteria itself is a callback function, so you can craft it to your needs.

When an email message is matched:
1. A folder will be created with the subject text prepended by date (YYYY.MM.DD);
2. Inside that folder, a text file will be created with the email message;
3. All attachments will be saved;

> Attention! This script only works on Windows AND Outlook desktop application must be installed. 


# How to use
## Requirements
Make sure you have pywin32 installed:
```shell
pip install -r requirements.txt
```


## How to run
Since this is a one time only script (for now, at least), i haven't created a command line.
So, to use, change the parameters:
```python
    main(
        target_email="my@email.com", # email address associated with the account you want to read
        attachment_base_folder=Path("attachments"),  # Base path to where you want to save the attachments
        email_filter_by_subject=partial(email_filter_by_subject_callback, # Callback function that will filter the messages by it's subject.
                                        target_keywords=[
                                            ["recibo", "pagamento"],
                                            ["informe", "rendimento"],
                                            ["imposto", "renda"],
                                            ["irpf", ]
                                        ]),
        verbose=True # if true, will print a bunch of informative messages.
    )
```
A couple notes:
- The callback is optional. if you pass None, all messages will be matched.
- I'm using partial because I wanted to make a bit more dynamic the way I analyzed the emails. In brazilian portuguese, those keywords represent common word combinations for emails with paycheck receipts and tax related messages.

# TODO
- Add command line
- Try to use parallel processing here.