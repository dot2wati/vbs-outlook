' ---------------------------------------------------------
' Outlook �����Կ��� �ش� ������ �̸��� URL ��������
' Outlook �� �α��ε� ���·� ������־�� ��
' LogFileName Parameter�� �Ѱ��ָ� append ���� UTF-8 ����
' ---------------------------------------------------------

' ������ UserPrifile ��� ���ϱ�
Set WshShell = WScript.CreateObject("WScript.Shell")
pathUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%") 

' ������ �Ķ���ʹ� pathLogFileName ���� ����
if wscript.arguments.count > 0 then
	lastArgNum = wscript.arguments.count - 1

end if

' �Ķ���� ���� Ȯ��
if wscript.arguments.count = 3 then 
	' outlook ����, ã�� ����, �α����ϰ��
	folderOutlook = WScript.Arguments(0)
	' ���Ͽ� ���Ե� �ؽ�Ʈ
	findText = WScript.Arguments(1)
	' ������ġ
	pathLogFileName = Wscript.Arguments(lastArgNum)

elseif wscript.arguments.count = 0 then
	' �Ķ���� ���� > �׽�Ʈ
	msgbox "getEmail.vbs > Test (�Ķ���� ����) > ����ȭ��\getEmailTest.log"
	' ���Ͽ� ���Ե� �ؽ�Ʈ
	findText = "MDG2004"
	' ������ġ
	folderOutlook = "RPA\MDG\MDG2004" 
	pathLogFileName = pathUserProfile & "\Desktop\getEmailTest.log"
	
else
	msgbox "�Ķ���� ���� �ȸ��� > ����"
	WScript.Quit

end if


' ���� vbs ��ġ ���� ���
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' �α׳���� vbs ���� ȣ�� 
pathVbsAppendLog = scriptdir & "\" & "appendLog.vbs"

' �α� �����
textLog = "���� VBScript"
WshShell.Run(pathVbsAppendLog + " " + pathLogFileName + " " + Chr(34) & textLog & Chr(34))


' outlookApp
Set outlookApp = CreateObject("Outlook.Application")
Set outlookMAPI = outlookApp.GetNameSpace("MAPI")


navFolder = Split(folderOutlook,"\")

' ���������� ����
' Set outlookFolder = outlookMAPI.GetDefaultFolder(6)

' ���������� ���� ����
Set outlookFolder = outlookMAPI.GetDefaultFolder(6).Parent
for each folderName in navFolder
	Set outlookFolder = outlookFolder.Folders(folderName)
next


' msgbox outlookFolder
Set allEmails = outlookFolder.Items
Dim grpMailBody

textLog = "���� �б� ����"
WshShell.Run(pathVbsAppendLog + " " + pathLogFileName + " " + Chr(34) & textLog & Chr(34))

For Each email In outlookFolder.Items
	if email.subject = findText Then
		grpMailBody = grpMailBody + "," + email.body
		
		' �α� ����� ȣ��
		textLog = email.body
		WshShell.Run(pathVbsAppendLog + " " + pathLogFileName + " " + Chr(34) & textLog & Chr(34))

		' �ش� ���� ����
		' email.Delete
		
	End IF
Next

' Regexp ù , ����
Set objReg=CreateObject("vbscript.regexp")
objReg.Pattern="^\s*,"
grpMailBody = objReg.Replace(grpMailBody,"")


IF wscript.arguments.count = 3 Then
	WScript.StdOut.WriteLine(grpMailBody)

else
	msgbox grpMailBody

End IF

'Quit
outlookApp.Quit

'Clear the memory
Set WshShell = Nothing
Set outlookApp = Nothing
Set outlookMAPI = Nothing
Set allEmails = Nothing
Set objReg = Nothing
