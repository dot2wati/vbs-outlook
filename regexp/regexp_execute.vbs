'Execute() �޼���� Test() �޼���� �޸� �������� Collection ��ü �������� ������ �ֹǷ� ��

Set objReg = CreateObject("vbscript.regexp")

'���������� ���ڵ�
objReg.Pattern = "[^\s]+"
objReg.Global = True

myText = "���ع��� ��λ��� ������ �⵵��"

Set myResults = objReg.Execute(myText)

msgbox myResults.Count
msgbox "Item(0):" & myResults.Item(0)

For Each res in myResults
    msgbox res
Next