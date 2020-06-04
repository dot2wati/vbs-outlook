' 첨부파일 저장


'function은 리턴값 사용가능, sub은 없음
sub appenText(pathMyFile, logText)
	' OpenTextFile 8이면 Append 옵션
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(pathMyFile,8,true)

	' 정규식 적용
	Set objReg=CreateObject("vbscript.regexp")
	objReg.Pattern = "[\s\r\n]+"
	objReg.Global = True
	logText = objReg.Replace(logText," ")

	writeData = "(" & now & ") " & logText

objFileToWrite.WriteLine(writeData)
	objFileToWrite.Close

End sub



' 윈도우 UserPrifile 경로 구하기
' 다운로드 위치 > 바탕화면 경로 (원하는 경로로 변경하여 사용가능)
Set WshShell = WScript.CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

pathUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%")
myDownPath = pathUserProfile & "\Desktop\downAttachFile"
logFilePathName = myDownPath & "\test.log"

' 폴더 생성
if oFSO.FolderExists(myDownPath) = False Then
	oFSO.CreateFolder myDownPath
end if


' 받은편지함
Set outlookApp = CreateObject("Outlook.Application")
Set outlookMAPI = outlookApp.GetNameSpace("MAPI")
Set outlookFolder = outlookMAPI.GetDefaultFolder(6)

' 메일위치
folderOutlook = "RPA\MDG\MDG2004" 
' Split
navFolder = Split(folderOutlook,"\")

' 폴더설정
Set outlookFolder = outlookMAPI.GetDefaultFolder(6).Parent
for each folderName in navFolder
	Set outlookFolder = outlookFolder.Folders(folderName)
next

For Each email In outlookFolder.Items

	appenText logFilePathName, "메일제목:" & email.Subject
	appenText logFilePathName, "메일수신:" & email.ReceivedTime

	For i = email.Attachments.Count to 1 step -1		
		appenText logFilePathName, "첨부파일:" & "(" & i & ")" & email.Attachments.Item(i).DisplayName
		email.Attachments.Item(i).SaveAsFile myDownPath & "\" & email.Attachments.Item(i).DisplayName
	Next

Next

Set WshShell = Nothing
Set oFSO = Nothing
Set outlookApp = Nothing
