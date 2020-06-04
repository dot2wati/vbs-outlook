' UTF-8 만들기
' 테스트해보니 UTF-8(BOM) 으로 만들어짐

Set objStream = CreateObject("ADODB.Stream")
objStream.CharSet = "utf-8"
objStream.Open

Set WshShell = CreateObject("wscript.Shell") 
pathUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%")
pathLogFileName = pathUserProfile & "\Desktop\test.log"

objStream.WriteText "(" & now & ") "& "The data I want in utf-8"
objStream.SaveToFile pathLogFileName, 2