' ------------------------------------------------------------------
' 이메일을 읽어서 원하는 경로의 csv 파일 만들기
' 제목,내용,보낸사람,시간
' ------------------------------------------------------------------

'function은 리턴값 사용가능, sub은 없음
sub appenText(pathMyFile, logText)
    ' OpenTextFile 8이면 Append 옵션
    Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(pathMyFile,8,true)

	' 정규식 적용
	Set objReg=CreateObject("vbscript.regexp")
	objReg.Pattern = "[\s\r\n]+"
	objReg.Global = True
	logText = objReg.Replace(logText," ")

    writeData = logText
    
    objFileToWrite.WriteLine(writeData)
	objFileToWrite.Close
	
End sub



' 윈도우 UserPrifile 경로 구하기
Set WshShell = WScript.CreateObject("WScript.Shell")
pathUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%")
' 결과파일 > 바탕화면 경로 (원하는 경로로 변경하여 사용가능)
pathTextFileName = pathUserProfile & "\Desktop\getEmailTest.csv"

' ------------------------------메일 읽기
' 메일에 구분 위한 텍스트
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

' 폴더의 메일 확인
' 여기서 해당 메일을 Delete 할 경우 전체 For Each 반복문 횟수도 줄어듦
' 오래된 메일 부터 수행함
For Each email In outlookFolder.Items
	' if InStr(email.subject,findText) 	' 텍스트 포함 조건 
	if email.subject = findText Then 	' 텍스트 일치 조건
		' Text 남기기 호출
		textBody = email.body
		appenText pathTextFileName, textBody

    End IF
Next

' 폴더의 조건에 해당하는 모든 메일 삭제
' step은 루프를 통해 매번 counter가 증가 하는 양입니다.
' 최근메일부터 수행 // 최근메일_num To 1 Step -1
For i = outlookFolder.Items.Count To 1 Step -1 
	If (outlookFolder.Items(i).Subject) = findText Then
		' outlookFolder.Items(i).Delete

	End If 
Next 


'Clear the memory
Set outlookMAPI = Nothing
Set outlookApp = Nothing
Set objFileToWrite = Nothing
Set objReg = Nothing