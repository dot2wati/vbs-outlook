if wscript.arguments.count > 0 then 
    pathLogFileName = WScript.Arguments(0)

elseif wscript.arguments.count = 0 then
    Set WshShell = CreateObject("wscript.Shell") 
    pathUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%")
    pathLogFileName = pathUserProfile & "\Desktop\test.log"
    
end if


Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(pathLogFileName,8,true)

writeData = Date & " " & test 
objFileToWrite.WriteLine(writeData)
objFileToWrite.Close
Set objFileToWrite = Nothing


