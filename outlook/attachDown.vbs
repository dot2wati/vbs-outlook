' ÷������ ����


'function�� ���ϰ� ��밡��, sub�� ����
sub appenText(pathMyFile, logText)
	' OpenTextFile 8�̸� Append �ɼ�
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(pathMyFile,8,true)

	' ���Խ� ����
	Set objReg=CreateObject("vbscript.regexp")
	objReg.Pattern = "[\s\r\n]+"
	objReg.Global = True
	logText = objReg.Replace(logText," ")

	writeData = "(" & now & ") " & logText

objFileToWrite.WriteLine(writeData)
	objFileToWrite.Close

End sub



' ������ UserPrifile ��� ���ϱ�
' �ٿ�ε� ��ġ > ����ȭ�� ��� (���ϴ� ��η� �����Ͽ� ��밡��)
Set WshShell = WScript.CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

pathUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%")
myDownPath = pathUserProfile & "\Desktop\downAttachFile"
logFilePathName = myDownPath & "\test.log"

' ���� ����
if oFSO.FolderExists(myDownPath) = False Then
	oFSO.CreateFolder myDownPath
end if


' ����������
Set outlookApp = CreateObject("Outlook.Application")
Set outlookMAPI = outlookApp.GetNameSpace("MAPI")
Set outlookFolder = outlookMAPI.GetDefaultFolder(6)

' ������ġ
folderOutlook = "RPA\MDG\MDG2004" 
' Split
navFolder = Split(folderOutlook,"\")

' ��������
Set outlookFolder = outlookMAPI.GetDefaultFolder(6).Parent
for each folderName in navFolder
	Set outlookFolder = outlookFolder.Folders(folderName)
next

For Each email In outlookFolder.Items

	appenText logFilePathName, "��������:" & email.Subject
	appenText logFilePathName, "���ϼ���:" & email.ReceivedTime

	For i = email.Attachments.Count to 1 step -1		
		appenText logFilePathName, "÷������:" & "(" & i & ")" & email.Attachments.Item(i).DisplayName
		email.Attachments.Item(i).SaveAsFile myDownPath & "\" & email.Attachments.Item(i).DisplayName
	Next

Next

Set WshShell = Nothing
Set oFSO = Nothing
Set outlookApp = Nothing
