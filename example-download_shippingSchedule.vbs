'================================================================================

'���������� ������ǰ ������ - ��������ǥ - yyyy-mm-dd ������ ���� ~ ������ ���������� ÷�������� �ٿ�ε� ��
'�ڡڡڡ� ���� �������Ϻ��� �Ϸ��� input �ϳ��� �޴°� ���� �ڡڡڡ�

'================================================================================
'�������� ��ġ ( ���� )
workPath = "\\150.2.80.150\rpa\TSM-001_������ǰ ������\��������ǥ\"

'�˻����� ( ���� )
text = "BOOKING"

Set olApp = CreateObject("Outlook.Application")
Set olMAPI = olApp.GetNameSpace("MAPI")

'������������ ���� ����
Set oFolder = olMAPI.GetDefaultFolder(6)
Set oFolder = oFolder.Folders("����������")

Set allEmails = oFolder.Items

'vDate = �Ϸ� �� ��¥

vDate = DateAdd("d",-1,Date)
vDate = clng(replace(vDate,"-",""))

For Each email In oFolder.Items
  	intCount  = email.Attachments.Count
	If intCount > 0 Then
		
		mailDate = mid(email.receivedtime,1,10)
		mailDate = clng(replace(mailDate,"-",""))
		
		if mailDate >= vDate then
			For i = 1 To intCount
				If InStr(email.subject,text) <> 0 then		
 					test = InStr(email.body,"Booking No")
					bookingNo =  trim(mid(email.body,test+11,12))
					'MsgBox email.subject & ", " &email.UnRead
					email.Unread = False
 					email.Attachments.Item(i).SaveAsFile workPath & "\" & bookingNo& ".pdf"
	 			End If
			Next
		else 
			
		end if

		
	end if
 Next
