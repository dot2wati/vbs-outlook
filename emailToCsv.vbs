' ------------------------------------------------------------------
' 이메일을 읽어서 원하는 경로의 csv 파일 만들기
' 제목,내용,보낸사람,시간
' ------------------------------------------------------------------

'function은 리턴값 사용가능, sub은 없음
sub appenText(logFileName, logText)
    ' OpenTextFile 8이면 Append 옵션
    Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(pathLogFileName,8,true)
    
    writeData = logText
    
    objFileToWrite.WriteLine(writeData)
    objFileToWrite.Close
End sub



' 윈도우 UserPrifile 경로 구하기
Set WshShell = WScript.CreateObject("WScript.Shell")
pathUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%")
' 결과파일 > 바탕화면 경로 (원하는 경로로 변경하여 사용가능)
pathLogFileName = pathUserProfile & "\Desktop\getEmailTest.csv"

' ------------------------------메일 읽기
' 메일에 포함된 텍스트
findText = "MDG2004"
' 메일위치
folderOutlook = "RPA\MDG\MDG2004" 
' Split
navFolder = Split(folderOutlook,"\")

Set outlookApp = CreateObject("Outlook.Application")
Set outlookMAPI = outlookApp.GetNameSpace("MAPI")


' 폴더설정
Set outlookFolder = outlookMAPI.GetDefaultFolder(6).Parent
for each folderName in navFolder
	Set outlookFolder = outlookFolder.Folders(folderName)
next

' 폴더의 메일 가져오기
Set allEmails = outlookFolder.Items
    
For Each email In outlookFolder.Items
	if email.subject = findText Then
		' 로그 남기기 호출
		textLog = email.body
		appenText pathLogFileName, textLog

		' 해당 메일 삭제
		' email.Delete
		
    End IF
    
Next


'Clear the memory
Set objFileToWrite = Nothing