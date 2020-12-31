import win32com.client as client
import re
import os

outlook = client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6) # "6" refers to the inbox folder. returns an object
narbutas_folder = inbox.Folders['Noreply Narbutas'] # refering / accessing the 'Leads' subfolder
messages = narbutas_folder.items

for message in messages:
	attachment = message.Attachments.Item(1) # select attachement item to save
	attachment.SaveAsFile(os.path.join(os.getcwd(), str(attachment))) #saves filse to current directory
	
# Add directory creatio where to save files
# Wrap in a function with arguments and optional filters
