'Execute() 메서드는 Test() 메서드와 달리 실행결과를 Collection 객체 형식으로 리턴해 주므로 굳

Set objReg = CreateObject("vbscript.regexp")

'공백제외한 문자들
objReg.Pattern = "[^\s]+"
objReg.Global = True

myText = "동해물과 백두산이 마르고 닳도록"

Set myResults = objReg.Execute(myText)

msgbox myResults.Count
msgbox "Item(0):" & myResults.Item(0)

For Each res in myResults
    msgbox res
Next