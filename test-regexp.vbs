' Test Regexp 
' http://blog.naver.com/PostView.nhn?blogId=jjjhyeok&logNo=20047243896&parentCategoryNo=33&viewDate=&currentPage=1&listtype=0

' ���Խ� ���� ����
Set objReg = CreateObject("vbscript.regexp")

objReg.Pattern = "\s*"
' vbs ���ԽĿ��� . (dot) ��ȣ�� ��� ������ �ٹٲ��� ��Ī��Ű�� ����
objReg.Pattern = "^.+"
' objReg.IgnoreCase = True
objReg.Global = True

myText = "��    �� " + vbCrLf + vbLf +"   ��"

msgbox myText

myText = objReg.Replace(myText,"")

msgbox myText

' Test2
Set objReg = CreateObject("vbscript.regexp")

objReg.Pattern = "^[^\s]+"
objReg.Global = True

myText = "���ع��� ��λ��� ������ �⵵��"

msgbox myText

' Replace�� ġȯ
' Test�� ���Կ��� �Ǵ� 
msgbox objReg.Test(myText)

' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/collection-object
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/item-method-visual-basic-for-applications
Set myResults = objReg.Execute(myText)
msgbox myResults.Count
msgbox myResults.Item(0)

For Each res in myResults
    msgbox res
Next