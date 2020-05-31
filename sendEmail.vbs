' ----------------------------------------------------------------------------------
' VBScript Encoding EUC-KR
' �Ķ���� 5�� ����, ����, To, CC, ÷������
' To, CC, ÷�������� ��� ;�� ������. ������ ���� ���� ����
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
    msgBox "�Ķ���� ���� �ٸ� > ����" & vbCrLf & "Prameter ���� > ����, ����, To, CC, ÷������" & vbCrLf & "To, CC, ÷�������� ;�� ����"
    Wscript.Quit

end if

' ���Խ� email ���� ����
Set objReg=CreateObject("vbscript.regexp")
objReg.Pattern="\s*"
myTo = objReg.Replace(myTo,"")
myCC = objReg.Replace(myCC,"")

' ���Խ� ù ���� ����
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
    ' ;�� �����Ͽ� ÷������ Add
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