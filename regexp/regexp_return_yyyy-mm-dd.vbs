' 년월일 받음

' 1. 일단 텍스트에서 [\\\.-년월]은 -로 치환 마지막 문자인데 숫자가 아니면 빈문자로 치환
' 2. \d{2}-\d{1,2}-\d{1,2}|\d{4}-\d{1,2}-\d{1,2}의 형태가 존재하면 년월일 존재
' 3. Test 메서드로 해당 정규식이 존재하는지 확인 가능
' 4. 존재하면 Execute 메서드로 년-월-일에 해당하는 부분 가져옴
' 5. 자리수 맞춤 년도는 4자리 월은 2자리 일도 2자리

' 이쁘게 Return yyyy-mm-dd 형태로 AA에서는 저형태 고정으로 가정하고 활용


' Execute() 메서드는 Test() 메서드와 달리 실행결과를 Collection 객체 형식으로 리턴해 주므로 굳

Set objReg = CreateObject("vbscript.regexp")

' 년월일 예시 
' 2020-05-08
' 20-5-8
' 년도가 2자리 4자리가능
' 월,일은 2자리 1자리 가능
objReg.Pattern = "^\d{2}-\d{1,2}-\d{1,2}$|^\d{4}-\d{1,2}-\d{1,2}$"
objReg.Global = True

' Test
' myDate = "  20200608  "
myDate = WScript.Arguments(0)

' 전달받은 Date 값 Trim
myDate = Trim(myDate)

' 해당 정규식을 포함하고 있는지 확인
checkPattern = objReg.Test(myDate)

' True / Flase
' msgbox checkPattern

' vbs 정규식은 후방탐색이 안됨
' Execute 매서드는 "Set 변수명"으로 결과 받음 매칭되는 모든 값을 반환함 (컬렉션 개체)
' Replace 매서드는 변수에 바로 할당하여 결과 받음
IF checkPattern = True Then
	
	' 년도 추출
	objReg.Pattern = "^\d{2}(?=-)|^\d{4}(?=-)"
	Set myYearCollection = objReg.Execute(myDate)
	myYear = myYearCollection.Item(0)
	' msgbox myYear
	
	' 월 추출
	objReg.Pattern = "^\d+-|-\d+$"
	myMonth = objReg.Replace(myDate,"")
	' msgbox myMonth
		
	' 일 추출
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
	
	' 값 리턴 및 Quit
	' MsgBox myDate
	WScript.StdOut.WriteLine(myDate)
	
	Set objReg = Nothing
	WScript.Quit
	
End If

' 숫자로만 6자리 혹은 8자리인 경우 탐색
objReg.Pattern = "^\d{6}$|^\d{8}$"

' 해당 정규식을 포함하고 있는지 확인
checkPattern = objReg.Test(myDate)

' Date 값이 6자리 혹은 8자리임
IF checkPattern = True Then

	' 6자리인 경우
	IF Len(myDate) = 6 Then
		myYear = Mid(myDate,1,2)
		myMonth = Mid(myDate,3,2)
		myDay = Mid(myDate,5,2)
		
		If Len(myYear) = 2 Then
		myYear = "20" & myYear
		End If
		
		myDate = myYear & "-" & myMonth & "-" & myDay
		
		' 값 리턴 및 Quit
		' MsgBox myDate
		WScript.StdOut.WriteLine(myDate)
		Set objReg = Nothing
		WScript.Quit
		
	End If
	
	' 8자리인 경우
	IF Len(myDate) = 8 Then
		myYear = Mid(myDate,1,4)
		myMonth = Mid(myDate,5,2)
		myDay = Mid(myDate,7,2)
		
		myDate = myYear & "-" & myMonth & "-" & myDay
		
		' 값 리턴 및 Quit
		' MsgBox myDate
		WScript.StdOut.WriteLine(myDate)
		Set objReg = Nothing
		WScript.Quit
		
	End If

End If

'해당사항없을 때
' 값 리턴 및 Quit
myDate = "-1"
WScript.StdOut.WriteLine(myDate)
Set objReg = Nothing
WScript.Quit
