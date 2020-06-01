' ---------------------------------------------------------
' Outlook 메일함에서 해당 제목의 이메일 URL 가져오기
' Outlook 이 로그인된 상태로 실행되있어야 함
' LogFileName Parameter로 넘겨주면 append 해줌 UTF-8 가능
' ---------------------------------------------------------

' 윈도우 UserPrifile 경로 구하기
Set WshShell = WScript.CreateObject("WScript.Shell")
pathUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%") 

' 마지막 파라미터는 pathLogFileName 으로 지정
if wscript.arguments.count > 0 then
	lastArgNum = wscript.arguments.count - 1

end if

' 파라미터 개수 확인
if wscript.arguments.count = 3 then 
	' outlook 폴더, 찾을 문자, 로그파일경로
	folderOutlook = WScript.Arguments(0)
	' 메일에 포함된 텍스트
	findText = WScript.Arguments(1)
	' 메일위치
	pathLogFileName = Wscript.Arguments(lastArgNum)

elseif wscript.arguments.count = 0 then
	' 파라미터 없음 > 테스트
	msgbox "getEmail.vbs > Test (파라미터 없음) > 바탕화면\getEmailTest.log"
	' 메일에 포함된 텍스트
	findText = "MDG2004"
	' 메일위치
	folderOutlook = "RPA\MDG\MDG2004" 
	pathLogFileName = pathUserProfile & "\Desktop\getEmailTest.log"
	
else
	msgbox "파라미터 개수 안맞음 > 종료"
	WScript.Quit

end if


' 현재 vbs 위치 폴더 경로
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' 로그남기는 vbs 파일 호출 
pathVbsAppendLog = scriptdir & "\" & "appendLog.vbs"

' 로그 남기기
textLog = "실행 VBScript"
WshShell.Run(pathVbsAppendLog + " " + pathLogFileName + " " + Chr(34) & textLog & Chr(34))


' outlookApp
Set outlookApp = CreateObject("Outlook.Application")
Set outlookMAPI = outlookApp.GetNameSpace("MAPI")


navFolder = Split(folderOutlook,"\")

' 받은편지함 폴더
' Set outlookFolder = outlookMAPI.GetDefaultFolder(6)

' 받은편지함 상위 폴더
Set outlookFolder = outlookMAPI.GetDefaultFolder(6).Parent
for each folderName in navFolder
	Set outlookFolder = outlookFolder.Folders(folderName)
next


' msgbox outlookFolder
Set allEmails = outlookFolder.Items
Dim grpMailBody

textLog = "메일 읽기 시작"
WshShell.Run(pathVbsAppendLog + " " + pathLogFileName + " " + Chr(34) & textLog & Chr(34))

For Each email In outlookFolder.Items
	if email.subject = findText Then
		grpMailBody = grpMailBody + "," + email.body
		
		' 로그 남기기 호출
		textLog = email.body
		WshShell.Run(pathVbsAppendLog + " " + pathLogFileName + " " + Chr(34) & textLog & Chr(34))

		' 해당 메일 삭제
		' email.Delete
		
	End IF
Next

' Regexp 첫 , 제거
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
