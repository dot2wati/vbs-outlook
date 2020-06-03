' Test Regexp 
' http://blog.naver.com/PostView.nhn?blogId=jjjhyeok&logNo=20047243896&parentCategoryNo=33&viewDate=&currentPage=1&listtype=0

' 정규식 공백 제거
Set objReg = CreateObject("vbscript.regexp")

objReg.Pattern = "\s*"
' vbs 정규식에서 . (dot) 기호는 모든 문자중 줄바꿈은 매칭시키지 않음
objReg.Pattern = "^.+"
' objReg.IgnoreCase = True
objReg.Global = True

myText = "김    영 " + vbCrLf + vbLf +"   덕"

msgbox myText

myText = objReg.Replace(myText,"")

msgbox myText

' Test2
Set objReg = CreateObject("vbscript.regexp")

objReg.Pattern = "^[^\s]+"
objReg.Global = True

myText = "동해물과 백두산이 마르고 닳도록"

msgbox myText

' Replace는 치환
' Test는 포함여부 판단 
msgbox objReg.Test(myText)

' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/collection-object
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/item-method-visual-basic-for-applications
Set myResults = objReg.Execute(myText)
msgbox myResults.Count
msgbox myResults.Item(0)

For Each res in myResults
    msgbox res
Next