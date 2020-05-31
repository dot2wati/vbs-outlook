' ----------------------------------------------------------------------------------
' VBScript Encoding EUC-KR
' 파라미터 5개 제목, 내용, To, CC, 첨부파일
' To, CC, 첨부파일은 모두 ;로 구분함. 여러개 인자 전달 가능
' ----------------------------------------------------------------------------------
Dim myOutlook, myMail    
Dim strFileText 
Dim objFileToRead    


if wscript.arguments.count = 5 then 
    mySubject = WScript.Arguments(0)
    myBody = WScript.Arguments(1)
    myTo = WScript.Arguments(2)
    myCC = WScript.Arguments(3)
    myAttachments = WScript.Arguments(4)

else
    msgBox "파라미터 개수 다름 > 종료" & vbCrLf & "Prameter 순서 > 제목, 내용, To, CC, 첨부파일" & vbCrLf & "To, CC, 첨부파일은 ;로 구분"
    Wscript.Quit

end if

' 정규식 email 공백 제거
Set objReg=CreateObject("vbscript.regexp")
objReg.Pattern="\s*"
myTo = objReg.Replace(myTo,"")
myCC = objReg.Replace(myCC,"")

' 정규식 첫 공백 제거
objReg.Pattern="^\s*"
myAttachments = objReg.Replace(myAttachments,"")

myFiles = Split(myAttachments,";")


Set myOutlook = CreateObject("Outlook.Application")    
Set myMail = myOutlook.CreateItem(0)    

With myMail    
    .Subject = mySubject   
    .Body = myBody
    .To = myTo
    .CC = myCC
    ' ;로 구분하여 첨부파일 Add
    for each myFile in myFiles
        .Attachments.Add myfile
    next
    .Send    
End With    

myOutlook.Quit

'Clear the memory
Set myOutlook = Nothing    
Set myMail = Nothing
Set objReg = Nothing 