' Text����� VBScript
' VBScript Encoding EUC-KR
' -------------------
'
' �ٸ� VBS���� ȣ���ϴ� ���
' WshShell ������Ʈ �����
' WshShell.Run ���� ���ϴ� VBS ȣ��
' �Ķ���ʹ� " " �����̽� 1ĭ ���� �����Ͽ� ���� ����
' �Ʒ��� ���� 
'
' Set WshShell = WScript.CreateObject("WScript.Shell")
' WshShell.Run(ȣ����vbs��� + " " + Text������ + " " + Text���泻�� )
'
' �ٸ� Automation Anywhere ���� RunScript ���� ȣ���� VBScript���� �ٽ� �ٸ� VBScript ������ ȣ���ϴ� ���� ������� ���� ( Error )
' ������ Command prompt ���� �����ϸ� ���������� �ٸ� ���ϵ� ȣ�� ���� Ȯ����..!
'
' Automation Anywhere�� Run Script Ŀ�ǵ带 ���� ����ϰ������
' �ϳ��� VBScript ���� �ȿ� Sub �̳� Function ���·� ����
' ����ϸ� ���� ���°� �� �� 


' Example
'
' �Ķ���� �ΰ��� �ִ� ���
if wscript.arguments.count = 2 then 
    pathLogFileName = WScript.Arguments(0)
    appendText = WScript.Arguments(1)

' �Ķ���� 0���ΰ�� ����ȭ�鿡 �׽�Ʈ Text ����
elseif wscript.arguments.count = 0 then
    msgBox "appendLog.vbs > Test > �Ķ���� 0���� ��� ����ȭ�鿡 Test Text ����"
    Set WshShell = CreateObject("wscript.Shell") 
    pathUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%")
    pathLogFileName = pathUserProfile & "\Desktop\test.log"
    appendText = "Test"
    
else 
    msgBox "�Ķ���� ���� 2���ʿ� 1.Text���� path, 2.Text���� ����"    
    WScript.Quit
    
end if

' OpenTextFile 8�̸� Append �ɼ���, https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(pathLogFileName,8,true)

writeData = "(" & Now & ") " & appendText 
objFileToWrite.WriteLine(writeData)
objFileToWrite.Close


'Clear the memory
Set objFileToWrite = Nothing




