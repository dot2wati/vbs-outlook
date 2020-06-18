' 시분 받음
' 시분 예시 
' 
' 형태는 hh:mm, h:m 처리 자리수는 각각 \d{1,2}:\d{1,2} 임
' 숫자 1,2자리인 경우 처리
'	ex) 10 > 10:00, 1 > 01:00
' 값이 비어있으면 23:59로 처리
' 리턴값은 시, 분이 명확해야함 hh:mm형태로 리턴 AA에서는 저형태로 가정하고 활용함

' 입력값
myTime = WScript.Arguments(0)

' 전달받은 myTime 값 Trim
myTime = Trim(myTime)

' 정규식 개체 사용
Set objReg = CreateObject("vbscript.regexp")
objReg.Pattern = "^\d{1,2}:\d{1,2}$"
objReg.Global = True

' 해당 정규식을 포함하고 있는지 확인
checkPattern = objReg.Test(myTime)

IF checkPattern = True Then
	
	' 시간 추출
	objReg.Pattern = ":\d{1,2}$"
	myHour = objReg.Replace(myTime,"")
	' msgbox myHour
	
	' 분 추출
	objReg.Pattern = "^\d{1,2}:"
	myMinute = objReg.Replace(myTime,"")
	' msgbox myMinute
	
	' myTime
	myTime = myHour & ":" & myMinute
	
	' 값 리턴 및 Quit
	' MsgBox myTime
	WScript.StdOut.WriteLine(myTime)
	Set objReg = Nothing
	WScript.Quit
	
End IF

' 숫자 1,2자리인 경우 처리
'	ex) 10 > 10:00, 1 > 01:00
objReg.Pattern = "^\d{1,2}$"

' 해당 정규식을 포함하고 있는지 확인
checkPattern = objReg.Test(myTime)

IF checkPattern = True Then
	
	IF Len(myTime) = 1 Then
		'hh:mm 형태로 변환
		myTime = "0" & myTime & ":" & "00"
		
		' 값 리턴 및 Quit
		WScript.StdOut.WriteLine(myTime)
		Set objReg = Nothing
		WScript.Quit
	End If
	
	IF Len(myTime) = 2 Then
		myTime = myTime & ":" & "00"
		
		' 값 리턴 및 Quit
		WScript.StdOut.WriteLine(myTime)
		Set objReg = Nothing
		WScript.Quit
	End If

End IF

'해당사항없을 때
' 값 리턴 및 Quit
myTime = "-1"
WScript.StdOut.WriteLine(myTime)
Set objReg = Nothing
WScript.Quit

