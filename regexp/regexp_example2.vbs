' �ú� ����
' �ú� ���� 
' 
' ���´� hh:mm, h:m ó�� �ڸ����� ���� \d{1,2}:\d{1,2} ��
' ���� 1,2�ڸ��� ��� ó��
'	ex) 10 > 10:00, 1 > 01:00
' ���� ��������� 23:59�� ó��
' ���ϰ��� ��, ���� ��Ȯ�ؾ��� hh:mm���·� ���� AA������ �����·� �����ϰ� Ȱ����

' �Է°�
myTime = WScript.Arguments(0)

' ���޹��� myTime �� Trim
myTime = Trim(myTime)

' ���Խ� ��ü ���
Set objReg = CreateObject("vbscript.regexp")
objReg.Pattern = "^\d{1,2}:\d{1,2}$"
objReg.Global = True

' �ش� ���Խ��� �����ϰ� �ִ��� Ȯ��
checkPattern = objReg.Test(myTime)

IF checkPattern = True Then
	
	' �ð� ����
	objReg.Pattern = ":\d{1,2}$"
	myHour = objReg.Replace(myTime,"")
	' msgbox myHour
	
	' �� ����
	objReg.Pattern = "^\d{1,2}:"
	myMinute = objReg.Replace(myTime,"")
	' msgbox myMinute
	
	' myTime
	myTime = myHour & ":" & myMinute
	
	' �� ���� �� Quit
	' MsgBox myTime
	WScript.StdOut.WriteLine(myTime)
	Set objReg = Nothing
	WScript.Quit
	
End IF

' ���� 1,2�ڸ��� ��� ó��
'	ex) 10 > 10:00, 1 > 01:00
objReg.Pattern = "^\d{1,2}$"

' �ش� ���Խ��� �����ϰ� �ִ��� Ȯ��
checkPattern = objReg.Test(myTime)

IF checkPattern = True Then
	
	IF Len(myTime) = 1 Then
		'hh:mm ���·� ��ȯ
		myTime = "0" & myTime & ":" & "00"
		
		' �� ���� �� Quit
		WScript.StdOut.WriteLine(myTime)
		Set objReg = Nothing
		WScript.Quit
	End If
	
	IF Len(myTime) = 2 Then
		myTime = myTime & ":" & "00"
		
		' �� ���� �� Quit
		WScript.StdOut.WriteLine(myTime)
		Set objReg = Nothing
		WScript.Quit
	End If

End IF

'�ش���׾��� ��
' �� ���� �� Quit
myTime = "-1"
WScript.StdOut.WriteLine(myTime)
Set objReg = Nothing
WScript.Quit

