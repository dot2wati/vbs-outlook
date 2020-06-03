'================================================================================

'공유폴더의 수출제품 출고업무 - 선적일정표 - yyyy-mm-dd 폴더에 전날 ~ 당일의 선적스케줄 첨부파일을 다운로드 함
'★★★★ 만약 전영업일부터 하려면 input 하나를 받는게 좋음 ★★★★

'================================================================================
'공유폴더 위치 ( 고정 )
workPath = "\\150.2.80.150\rpa\TSM-001_수출제품 출고업무\선적일정표\"

'검색조건 ( 고정 )
text = "BOOKING"

Set olApp = CreateObject("Outlook.Application")
Set olMAPI = olApp.GetNameSpace("MAPI")

'보낸편지함의 내부 폴더
Set oFolder = olMAPI.GetDefaultFolder(6)
Set oFolder = oFolder.Folders("선적스케줄")

Set allEmails = oFolder.Items

'vDate = 하루 전 날짜

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
