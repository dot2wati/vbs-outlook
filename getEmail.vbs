'==================================================


'==================================================


' �˻����� ( ���� )
strUser = CreateObject("WScript.Network").UserName
' MsgBox strUser

Set WshShell = CreateObject("wscript.Shell") 
pathUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%") 
' MsgBox pathUserProfile

findText = "MDG2004"

Set olApp = CreateObject("Outlook.Application")
Set olMAPI = olApp.GetNameSpace("MAPI")
' olMAPI.Logon
' olMAPI.SendAndReceive(True) 


' ���������� ����
Set oFolder = olMAPI.GetDefaultFolder(6)
' msgbox oFolder
' ���������� ���� ����
' Set oFolder = oFolder.Folders(".���� �� ������")
' msgbox oFolder
Set allEmails = oFolder.Items


' vDate = �Ϸ� �� ��¥
' Date �� ���� ��¥
vDate = DateAdd("d",-1,Date)

'DateAdd("h",1,"31-Jan-10 08:50:00")
nowDate = Now()

calcDateTime = DateAdd("h",-12,nowDate)

' MsgBox nowDate & "  // " & calcDateTime

' MsgBox vDate
' MsgBox vDate > Date
' MsgBox vDate > nowDate

' WScript.Quit

'vDate = clng(replace(vDate,"-",""))

d=CDate(nowDate)
' msgbox d

For Each email In oFolder.Items
	' MsgBox email.receivedtime & " // " & calcDateTime
	' MsgBox email.receivedtime > calcDateTime
	' MsgBox email.subject
	intCount  = email.Attachments.Count
	
	if email.subject = findText Then
		
		If email.receivedtime > calcDateTime Then
			MsgBox email.body
		End IF
	End IF
	' MsgBox email
	' Set emailTime = email.ReceivedTime
	
	' if emailTime > calcDateTime Then
		
	' Else
		
	' End IF

	' Exit For
Next
