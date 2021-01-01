import win32com.client as client
import os

outlook = client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6) # "6" refers to the inbox folder. returns an object
index_folder = inbox.Folders['Folder name here'] # refering / accessing the needed subfolder in the inbox
messages = index_folder.items
folder_name = str(index_folder) # Creating a folder name

try:
	os.mkdir(os.path.join(os.getcwd(), folder_name)) # creating new directory (os.mkdir) by joining(os.path.join) current directory path(os.getcwd()) with new folder name
except:
	pass
os.chdir(os.path.join(os.getcwd(), folder_name)) # change to the new directory
for message in messages:
	attachment = message.Attachments.Item(1) # select attachement item to save
	# attachment.SaveAsFile(os.path.join(os.getcwd(), str(attachment))) # saves filse to current directory
	attachment_1 = message.Attachments.Item(2) # select attachement item to save
	# attachment.SaveAsFile(os.path.join(os.getcwd(), str(attachment_1))) # saves filse to current directory
	print(attachment_1)
	print(attachment)

# Add download all attachement files in emails functional
# Add code to check whether the file is already downloaded
# Wrap in a function with arguments and optional filters
