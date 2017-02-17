Dim objShell, objFso

Set objShell = CreateObject("Wscript.Shell")

SYSTEM_VBOX_APP_PATH=objShell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Oracle\VirtualBox\VBoxManage.exe"
SYSTEM_DOCKERMACHINE_APP_PATH=objShell.ExpandEnvironmentStrings("%DOCKER_TOOLBOX_INSTALL_PATH%") & "\docker-machine.exe"

VBOX_APP_PATH=""
DOCKERMACHINE_APP_PATH=""

Set objFso = CreateObject("Scripting.FileSystemObject")
If (objFso.FileExists(SYSTEM_VBOX_APP_PATH)) Then
    VBOX_APP_PATH = SYSTEM_VBOX_APP_PATH
End If

If (objFso.FileExists(SYSTEM_DOCKERMACHINE_APP_PATH)) Then
    DOCKERMACHINE_APP_PATH = SYSTEM_DOCKERMACHINE_APP_PATH
End If

If (DOCKERMACHINE_APP_PATH <> "") Then
    objShell.Run """" & DOCKERMACHINE_APP_PATH & """" & " create --driver virtualbox default", 2
End If

If (VBOX_APP_PATH <> "") Then
    objShell.Run """" & VBOX_APP_PATH & """" & " startvm ""default"" --type headless", 2
End If

Set USER_ENVIRONMENT = objShell.Environment("USER")

If (DOCKERMACHINE_APP_PATH <> "") Then
    Set objExecObject = objShell.Exec("""" & DOCKERMACHINE_APP_PATH & """" & " env")
    Do Until objExecObject.StdOut.AtEndOfStream
        strLine = objExecObject.StdOut.ReadLine()
        strSet = Instr(strLine,"SET")
        If strSet <> 0 Then
            firstSplit=Split(strLine)
            
            varAndSetting=Split(firstSplit(1), "=")
            USER_ENVIRONMENT( varAndSetting(0) ) = varAndSetting(1)
        End If
    Loop
End If