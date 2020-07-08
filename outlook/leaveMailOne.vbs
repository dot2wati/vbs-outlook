' 인수 받기 
' AA로 vbs 호출시 해당 호출 파일을 파라미터의 마지막 값으로 전달함
' 경로는 엄청 긺

' 내가 만약 수동으로 파라미터를 넘긴다면
' ex) 1 2 3 4 5 6 10
' 첫번 째 파라미터는 2자리 이내의 숫자라고 가정
' 1. 파라미터의 Length가 2보다 크면 파라미터를 전달받지 못한 것으로 간주
' 2. 파라미터의 값은 2자리 이내의 숫자로 가정
' 3. 에러시 리턴값은 -1

' 두번 째 파라미터는 다운로드 경로임
' 테스트 하는 중에는 바탕화면에 폴더 다운로드

' 임의로 남길 메일 수 지정 Default 값 1
countLeaveMail = 3

' 받은편지함
Set outlookApp = CreateObject("Outlook.Application")
Set outlookMAPI = outlookApp.GetNameSpace("MAPI")
Set outlookFolder = outlookMAPI.GetDefaultFolder(6)

' 메일위치 변수로 받을 예정
folderOutlook = "RPA\MDG\MDG2008" 
' Split
navFolder = Split(folderOutlook,"\")

' 폴더설정
Set outlookFolder = outlookMAPI.GetDefaultFolder(6).Parent
for each folderName in navFolder
	Set outlookFolder = outlookFolder.Folders(folderName)
next

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

mailCount =  outlookFolder.Items.Count
mailCountLimit = outlookFolder.Items.Count - countLeaveMail

' 폴더의 조건에 해당하는 모든 메일 삭제
' step은 루프를 통해 매번 counter가 증가 하는 양
' 최근메일부터 수행 // 최근메일_num To 1 Step -1
For i = outlookFolder.Items.Count To 1 Step -1 
    If i <= mailCountLimit Then 
        outlookFolder.Items(i).Delete
    End If
Next 

' 다운로드 flag선언
flagFileDown = 0

' 첨부파일다운로드 진행
' 삭제 후에 남은 메일에서 가장 최근메일 하나에서 첨부파일 다운로드
For i = outlookFolder.Items.Count To 1 Step -1 

    ' 첨부파일 개수가 1 미만 다운로드 시도하면 Error
    ' 첨부파일 개수가 1 이상 다운로드 시도
    IF outlookFolder.Items(i).Attachments.Count >= 1 Then
        For j = outlookFolder.Items(i).Attachments.Count to 1 step -1
            outlookFolder.Items(i).Attachments.Item(j).SaveAsFile myDownPath & "\" & outlookFolder.Items(i).Attachments.Item(j).DisplayName
            ' 한번이라도 다운받으면 flag값 1
            flagFileDown = 1
        Next
    End IF

    ' 가장 최근 파일 하나만 받을 것이므로 다운로드 후 
    ' 특정 조건 없이 Exit For
    Exit For
Next 

' 값 flagFileDown 리턴하여 다운로드 정상 수행되었는지 확인