Set Shell = CreateObject("Shell.Application")
Set WShell = WScript.CreateObject("WScript.Shell")
Set ProcEnv = WShell.Environment("PROCESS")

cmd = ProcEnv("CMD")

If (WScript.Arguments.Count >= 1) Then
  Shell.ShellExecute "cmd", "/c ""%cmd%""", "", "runas", 0
Else
  WScript.Quit
End If
