import win32com.client
import win32com
import os
import sys
import pyttsx3

speaker = pyttsx3.init()
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
message = messages.GetLast()
sender = message.Sender
if sender != "":
    speaker.say('You have mail from :'+sender.name)
speaker.runAndWait()






