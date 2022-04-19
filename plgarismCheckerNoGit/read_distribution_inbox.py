# Get E-Mail Id's from the particular distribution List

import win32com.client

# Distribution List Name
DL_NAME = 'dummy_group'

# Outlook
outApp = win32com.client.gencache.EnsureDispatch("Outlook.Application")
# print("OUTAPP", outApp)

# Get contact folder
contact_folder = outApp.Session.GetDefaultFolder(win32com.client.constants.olFolderContacts)
print("CONTACT FOLDER", contact_folder.Items)
print(win32com.client.constants.olDistributionList)
print(contact_folder.Items.Count)
# Iterate through contacts
for contact_item in contact_folder.Items:
    print("CONTACT_ITEM", contact_item)
    # check if item is a distribution list
    if contact_item.Class == win32com.client.constants.olDistributionList:
        print("Enter Distribution List")
        # check if distribution list's name is equal to the constant
        if contact_item.DLName == DL_NAME:
            print("Distribution List Name", contact_item.DLName)
            # loop through distribution list members and get their email address
            for i in range(1, contact_item.MemberCount + 1):
                member_in_dl = contact_item.GetMember(i)
                print('{} | {}'.format(contact_item.DLName, member_in_dl.Address))
    else:
        print("Enter else if")
    print("OK")
