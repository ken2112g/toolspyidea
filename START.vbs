Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = strPath

' Chạy Python ẩn (không hiện terminal)
WshShell.Run "cmd /c cd /d """ & strPath & """ && python main.py", 0, False

' Đợi server start rồi mở trình duyệt
WScript.Sleep 2500
WshShell.Run "http://127.0.0.1:5123", 1, False
