' ����� ����

' 1. �ϴ� �ؽ�Ʈ���� [\\\.-���]�� -�� ġȯ ������ �����ε� ���ڰ� �ƴϸ� ���ڷ� ġȯ
' 2. \d{2}-\d{1,2}-\d{1,2}|\d{4}-\d{1,2}-\d{1,2}�� ���°� �����ϸ� ����� ����
' 3. Test �޼���� �ش� ���Խ��� �����ϴ��� Ȯ�� ����
' 4. �����ϸ� Execute �޼���� ��-��-�Ͽ� �ش��ϴ� �κ� ������
' 5. �ڸ��� ���� �⵵�� 4�ڸ� ���� 2�ڸ� �ϵ� 2�ڸ�

' �̻ڰ� Return yyyy-mm-dd ���·� AA������ ������ �������� �����ϰ� Ȱ��


' Execute() �޼���� Test() �޼���� �޸� �������� Collection ��ü �������� ������ �ֹǷ� ��

Set objReg = CreateObject("vbscript.regexp")

' ����� ���� 
' 2020-05-08
' 20-5-8
' �⵵�� 2�ڸ� 4�ڸ�����
' ��,���� 2�ڸ� 1�ڸ� ����
objReg.Pattern = "^\d{2}-\d{1,2}-\d{1,2}$|^\d{4}-\d{1,2}-\d{1,2}$"
objReg.Global = True

' Test
' myDate = "  20200608  "
myDate = WScript.Arguments(0)

' ���޹��� Date �� Trim
myDate = Trim(myDate)

' �ش� ���Խ��� �����ϰ� �ִ��� Ȯ��
checkPattern = objReg.Test(myDate)

' True / Flase
' msgbox checkPattern

' vbs ���Խ��� �Ĺ�Ž���� �ȵ�
' Execute �ż���� "Set ������"���� ��� ���� ��Ī�Ǵ� ��� ���� ��ȯ�� (�÷��� ��ü)
' Replace �ż���� ������ �ٷ� �Ҵ��Ͽ� ��� ����
IF checkPattern = True Then
	
	' �⵵ ����
	objReg.Pattern = "^\d{2}(?=-)|^\d{4}(?=-)"
	Set myYearCollection = objReg.Execute(myDate)
	myYear = myYearCollection.Item(0)
	' msgbox myYear
	
	' �� ����
	objReg.Pattern = "^\d+-|-\d+$"
	myMonth = objReg.Replace(myDate,"")
	' msgbox myMonth
		
	' �� ����
	objReg.Pattern = "^\d+-\d+-"
	myDay = objReg.Replace(myDate,"")
	' msgbox myMonth
	
	If Len(myYear) = 2 Then
		myYear = "20" & myYear
	End If
	
	If Len(myMonth) = 1 Then
		myMonth = "0" & myMonth
	End If
	
	If Len(myDay) = 1 Then
		myDay = "0" & myDay
	End If
	
	myDate = myYear & "-" & myMonth & "-" & myDay
	
	' �� ���� �� Quit
	' MsgBox myDate
	WScript.StdOut.WriteLine(myDate)
	
	Set objReg = Nothing
	WScript.Quit
	
End If

' ���ڷθ� 6�ڸ� Ȥ�� 8�ڸ��� ��� Ž��
objReg.Pattern = "^\d{6}$|^\d{8}$"

' �ش� ���Խ��� �����ϰ� �ִ��� Ȯ��
checkPattern = objReg.Test(myDate)

' Date ���� 6�ڸ� Ȥ�� 8�ڸ���
IF checkPattern = True Then

	' 6�ڸ��� ���
	IF Len(myDate) = 6 Then
		myYear = Mid(myDate,1,2)
		myMonth = Mid(myDate,3,2)
		myDay = Mid(myDate,5,2)
		
		If Len(myYear) = 2 Then
		myYear = "20" & myYear
		End If
		
		myDate = myYear & "-" & myMonth & "-" & myDay
		
		' �� ���� �� Quit
		' MsgBox myDate
		WScript.StdOut.WriteLine(myDate)
		Set objReg = Nothing
		WScript.Quit
		
	End If
	
	' 8�ڸ��� ���
	IF Len(myDate) = 8 Then
		myYear = Mid(myDate,1,4)
		myMonth = Mid(myDate,5,2)
		myDay = Mid(myDate,7,2)
		
		myDate = myYear & "-" & myMonth & "-" & myDay
		
		' �� ���� �� Quit
		' MsgBox myDate
		WScript.StdOut.WriteLine(myDate)
		Set objReg = Nothing
		WScript.Quit
		
	End If

End If

'�ش���׾��� ��
' �� ���� �� Quit
myTime = "-1"
WScript.StdOut.WriteLine(myDate)
Set objReg = Nothing
WScript.Quit