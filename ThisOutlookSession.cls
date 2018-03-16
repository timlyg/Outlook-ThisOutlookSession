VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'' Move Outlook to top when new mail arrives
'Private Sub Application_NewMail()
'    Outlook.Application.ActiveWindow.Activate
'End Sub



'https://www.tachytelic.net/2017/10/how-to-run-a-vba-macro-when-new-mail-is-received-in-outlook/
Option Explicit
Private WithEvents inboxItems As Outlook.Items
Attribute inboxItems.VB_VarHelpID = -1

' declare this object withEvents displaying all the events
Private WithEvents olRemind As Outlook.Reminders
Attribute olRemind.VB_VarHelpID = -1
Dim ReminderAccount As String


Private Sub Application_Startup()
  Dim outlookApp As Outlook.Application
  Dim objectNS As Outlook.NameSpace
  
  Set outlookApp = Outlook.Application
  Set objectNS = outlookApp.GetNamespace("MAPI")
  Set inboxItems = objectNS.GetDefaultFolder(olFolderInbox).Items
End Sub


Private Sub inboxItems_ItemAdd(ByVal Item As Object)
On Error GoTo ErrorHandler
Dim Msg As Outlook.MailItem
Dim MessageInfo
Dim Result
If TypeName(Item) = "MailItem" Then
'    MessageInfo = "" & _
'        "Sender : " & Item.SenderEmailAddress & vbCrLf & _
'        "Sent : " & Item.SentOn & vbCrLf & _
'        "Received : " & Item.ReceivedTime & vbCrLf & _
'        "Subject : " & Item.Subject & vbCrLf & _
'        "Size : " & Item.Size & vbCrLf & _
'        "Message Body : " & vbCrLf & Item.Body
'    Result = MsgBox(MessageInfo, vbOKOnly, "New Message Received")

    Outlook.Application.ActiveWindow.Activate

End If
ExitNewItem:
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " - " & Err.Description
    Resume ExitNewItem
End Sub



' Dismisses Reminders based on Accounts, before they appear. This works together with Private WithEvents olRemind from top of the code.
Private Sub Application_Reminder(ByVal Item As Object)
    Set olRemind = Outlook.Reminders
    ReminderAccount = Item.Parent.Parent
End Sub

Private Sub olRemind_BeforeReminderShow(Cancel As Boolean)
    Dim objRem As Reminder
    For Each objRem In olRemind
                If objRem.IsVisible And ReminderAccount = "TLaw@solomonpage.com" Then
                    MsgBox ReminderAccount
                    objRem.Dismiss  ' dismisses reminder
                    Cancel = True   ' takes out the reminder window
                End If
                Exit For
        Next objRem

End Sub
