<div align="center">

## Get Outlook 2002 Contact Info


</div>

### Description

I wrote a program that gets contacts from Outlook, although it worked in Outlook 2000, it did not work in 2002. This code also works in 2002, without the Outlook Object 10.0. You can use 2000's 9.0 Object Library and it still works. This caused me so much trouble, I hope it helps someone else. Please leave comments or vote if it helps.
 
### More Info
 
Assumes you have added Microsoft Outlook Object Library 9.0 or 10.0 as a reference. And your form contains a listbox named list1.

Not sure if it works with VB 5.0.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Erica Ziegler\-Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/erica-ziegler-roberts.md)
**Level**          |Intermediate
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/erica-ziegler-roberts-get-outlook-2002-contact-info__1-45371/archive/master.zip)

### API Declarations

```
Dim contArray()
Dim ol as Outlook.Application
```


### Source Code

```
Private Sub Form_Load()
  Dim olns As NameSpace
  Dim itemCount As Integer
  Dim objfolder As mapiFolder
  Dim objAllContacts As Outlook.Items
  Dim i As Variant
  Dim Contact As Outlook.ContactItem
  ReDim contArray(3, 50)
  Me.restore.Enabled = False
  Me.minimize.Enabled = True
  'Create an instance of Outlook
  Set ol = CreateObject("Outlook.Application")
  Set olns = ol.GetNamespace("MAPI")
  olns.Logon
  Set objfolder = olns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts)
  Set objAllContacts = objfolder.Items
  itemCount = objAllContacts.Count
  List1.Clear
  i = 0
  For i = 1 To itemCount
    If TypeOf objAllContacts.Item(i) Is Outlook.ContactItem Then
      Set Contact = objAllContacts.Item(i)
      If Contact.CompanyName <> "" Then
        contArray(1, i) = Contact.CompanyName
        contArray(2, i) = Contact.BusinessTelephoneNumber
        contArray(3, i) = Contact.BusinessFaxNumber
        List1.AddItem Contact.CompanyName
      End If
      If i = UBound(contArray, 2) Then
        ReDim Preserve contArray(3, i + 50)
      End If
    End If
      'i = i + 1
  Next
  olns.Logoff
  Set olns = Nothing
  Set objfolder = Nothing
  Set objAllContacts = Nothing
  Set Contact = Nothing
End Sub
```

