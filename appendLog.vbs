' ----------------------------------------------------------------------------------
' �α׳���� VBScript
' VBScript Encoding EUC-KR
' ----------------------------------------------------------------------------------
' �ٸ� VBS���� ȣ���ϴ� ���
' WshShell ������Ʈ �����
' WshShell.Run ���� ���ϴ� VBS ȣ��
' �Ķ���ʹ� " " �����̽� 1ĭ ���� �����Ͽ� ���� ����
' �Ʒ��� ���� 

' Set WshShell = WScript.CreateObject("WScript.Shell")
' WshShell.Run(ȣ����vbs��� + " " + �α׳����� + " " + �α׳��泻�� )
' ----------------------------------------------------------------------------------

' �Ķ���� �ΰ��� �ִ� ���
if wscript.arguments.count = 2 then 
    pathLogFileName = WScript.Arguments(0)
    appendText = WScript.Arguments(1)

' �Ķ���� 0���ΰ�� ����ȭ�鿡 �׽�Ʈ �α� ����
elseif wscript.arguments.count = 0 then
    msgBox "appendLog.vbs > Test > �Ķ���� 0���� ��� ����ȭ�鿡 Test �α� ����"
    Set WshShell = CreateObject("wscript.Shell") 
    pathUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%")
    pathLogFileName = pathUserProfile & "\Desktop\test.log"
    appendText = "Test"
    
else 
    msgBox "�Ķ���� ���� 2���ʿ� 1.�α����� path, 2.�α׳��� ����"    
    WScript.Quit
    
end if

' OpenTextFile 8�̸� Append �ɼ���, https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(pathLogFileName,8,true)

writeData = "(" & Now & ") " & appendText 
objFileToWrite.WriteLine(writeData)
objFileToWrite.Close


'Clear the memory
Set objFileToWrite = Nothing




