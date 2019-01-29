Sub Reply()
Dim oMItem As Outlook.MailItem
Dim oMItemReply As Outlook.MailItem
Dim sGreetName As String
Dim komma As Integer
Dim kompletterName, Name, Vorname As String


On Error Resume Next
Select Case TypeName(Application.ActiveWindow)
Case "Explorer"
Set oMItem = ActiveExplorer.Selection.Item(1)
Case "Inspector"
Set oMItem = ActiveInspector.CurrentItem
Case Else
End Select
On Error GoTo 0
If oMItem Is Nothing Then GoTo ExitProc
On Error Resume Next
sGreetName = oMItem.SenderName

Set oMItemReply = oMItem.Reply

sGreetName = Trim(Mid(sGreetName, 1, InStr(sGreetName, " ") - 1))

sGreetName = Replace(sGreetName, ",", "")

komma = InStr(sGreetName, ",")
Name = Mid(sGreetName, 1, komma - 1)
Vorname = Mid(sGreetName, komma + 1, Len(sGreetName))

With oMItemReply
.HTMLBody = "<span style=""font-family:Arial;font-size : 10pt""><p>Hallo " & Vorname & "</p><p></p></span>" & .HTMLBody
.Display
End With

ExitProc:
Set oMItem = Nothing
Set oMItemReply = Nothing
SendKeys "{DOWN}", True
End Sub
