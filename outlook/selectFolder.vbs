' ---------------------------------------------------------------
' Outlook이 로그인된 상태로 실행 되어있는 상태여야 함
' 메일 폴더 선택
' ---------------------------------------------------------------

' 메일위치
folderOutlook = "RPA\MDG\MDG2004"
' Split
navFolder = Split(folderOutlook,"\")
' outlookApp
Set OutlookApp = CreateObject("Outlook.Application")
Set outlookMAPI = outlookApp.GetNameSpace("MAPI")

' 폴더설정
Set outlookFolder = outlookMAPI.GetDefaultFolder(6).Parent
for each folderName in navFolder
    Set outlookFolder = outlookFolder.Folders(folderName)
next

outlookApp.ActiveExplorer.SelectFolder outlookFolder

'Clear the memory
Set outlookApp = Nothing
Set outlookMAPI = Nothing
Set outlookFolder = Nothing