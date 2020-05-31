' ---------------------------------------------------------
' Outlook 메일함에서 해당 제목의 이메일 URL 가져오기
' Outlook 이 로그인된 상태로 실행되있어야 함
' LogFileName Parameter로 넘겨주면 append 해줌 UTF-8 가능
' ---------------------------------------------------------

' 윈도우 UserPrifile 경로 구하기
Set WshShell = WScript.CreateObject("WScript.Shell")
pathUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%") 

' 마지막 파라미터는 pathLogFileName으로 지정
if wscript.arguments.count > 0 then
	lastArgNum = wscript.arguments.count - 1
	
end if

' 파라미터 개수 확인
if wscript.arguments.count = 3 then 
	'outlook 폴더, 찾을 문자, 로그파일경로
	folderOutlook = WScript.Arguments(0)
	' 메일에 포함된 텍스트
	findText = WScript.Arguments(1)
	' 메일위치
	pathLogFileName = Wscript.Arguments(lastArgNum)

elseif wscript.arguments.count = 0 then
	'파라미터 없음 테스트
	msgbox "파라미터 개수 0 > Test > 바탕화면\getEmailTest.log"
	' 메일에 포함된 텍스트
	findText = "MDG2004"
	' 메일위치
	folderOutlook = "RPA\MDG2004" 
	pathLogFileName = pathUserProfile & "\Desktop\getEmailTest.log"
	
else
	msgbox "파라미터 개수 안맞음 > 종료"
	WScript.Quit

end if


' 현재 vbs 위치 폴더 경로
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' 로그남기는 vbs 파일 호출 
pathVbsAppendLog = scriptdir & "\" & "appendLog.vbs"

' 공백이 파라미터가 되므로 공백 제거함
textLog = "실행 VBScript"
textLog = Replace(textLog," ","_")

' 호출
WshShell.Run(pathVbsAppendLog + " " + pathLogFileName + " " + textLog)


'outlookApp
Set outlookApp = CreateObject("Outlook.Application")
Set outlookMAPI = outlookApp.GetNameSpace("MAPI")

' 받은편지함 폴더 > RPA > MDG2004
navFolder = Split(folderOutlook,"\")
Set outlookFolder = outlookMAPI.GetDefaultFolder(6)
for each folderName in navFolder
	Set outlookFolder = outlookFolder.Folders(folderName)
next

' msgbox outlookFolder
Set allEmails = outlookFolder.Items


' vDate = 하루 전 날짜
' Date 는 현재 날짜
vDate = DateAdd("d",-1,Date)
nowDate = Now()
calcDateTime = DateAdd("h",-12,nowDate)

' MsgBox nowDate & "  // " & calcDateTime
' MsgBox vDate
' MsgBox vDate > Date
' MsgBox vDate > nowDate

'vDate = clng(replace(vDate,"-",""))

d=CDate(nowDate)
' msgbox d

For Each email In outlookFolder.Items
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
