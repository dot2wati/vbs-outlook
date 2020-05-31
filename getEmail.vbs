' ---------------------------------------------------------
' Outlook �����Կ��� �ش� ������ �̸��� URL ��������
' Outlook �� �α��ε� ���·� ������־�� ��
' LogFileName Parameter�� �Ѱ��ָ� append ���� UTF-8 ����
' ---------------------------------------------------------

' ������ UserPrifile ��� ���ϱ�
Set WshShell = WScript.CreateObject("WScript.Shell")
pathUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%") 

' ������ �Ķ���ʹ� pathLogFileName���� ����
if wscript.arguments.count > 0 then
	lastArgNum = wscript.arguments.count - 1
	
end if

' �Ķ���� ���� Ȯ��
if wscript.arguments.count = 3 then 
	'outlook ����, ã�� ����, �α����ϰ��
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
	folderOutlook = "RPA\MDG2004" 
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
' ������ �Ķ���Ͱ� �ǹǷ� ���� ������
textLog = "���� VBScript"
textLog = Replace(textLog," ","_")

' ȣ��
WshShell.Run(pathVbsAppendLog + " " + pathLogFileName + " " + textLog)


' outlookApp
Set outlookApp = CreateObject("Outlook.Application")
Set outlookMAPI = outlookApp.GetNameSpace("MAPI")

' ���������� ���� > RPA > MDG2004
navFolder = Split(folderOutlook,"\")
Set outlookFolder = outlookMAPI.GetDefaultFolder(6)
for each folderName in navFolder
	Set outlookFolder = outlookFolder.Folders(folderName)
next

' msgbox outlookFolder
Set allEmails = outlookFolder.Items

Dim grpMailBody

For Each email In outlookFolder.Items
	if email.subject = findText Then
		grpMailBody = grpMailBody + "," email.body
		
		' �α� ����� ȣ��
		textLog = email.body
		WshShell.Run(pathVbsAppendLog + " " + pathLogFileName + " " + Chr(34) & textLog & Chr(34))

	End IF
Next

IF wscript.arguments.count = 3 Then 
	WScript.StdOut.Write(grpMailBody)
else
	msgbox grpMailBody
End IF