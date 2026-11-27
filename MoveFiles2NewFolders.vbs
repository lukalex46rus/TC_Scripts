'===============================================================================
' Script:      MoveFiles2NewFolders.vbs
' Version:     4.0
' Purpose:     Переместить выделенные элементы в новую/указанную папку
' Description: Если имя папки не задано — спрашивает. Создаёт папку при
'              необходимости. Перемещает каждый элемент из списка (%L).
'              Использует FSO.MoveFile/MoveFolder (без Shell.Application).
'
' SYNTAX:
'   cscript //nologo MoveFiles2NewFolders.vbs <ListFile> <Path> [Folder] [/Jump] [/Quiet] [/NoLog]
'
' PARAMETERS:
'   <ListFile> - %L
'   <Path>     - %P или %T
'   [Folder]   - имя папки; если пусто — InputBox
'
' KEYS:
'   /Jump   - после операции перейти в целевую папку
'   /Quiet  - без диалогов (требует [Folder])
'   /NoLog  - отключить лог
'
' RETURN CODES: 0 успех, 1 ошибка, 2 аргументы
' DEPENDENCIES: FSO, WScript.Shell
'===============================================================================
Option Explicit

Const SCRIPT_NAME = "MoveFiles2NewFolders"
Dim g_LogEnabled : g_LogEnabled = True
Dim g_LogFile    : g_LogFile = ""
Dim g_StartTime  : g_StartTime = Now

'--- ARGS ---
Dim oArgs, sList, sPath, sFolder, bJump, bQuiet, i, sA
Set oArgs = WScript.Arguments
If oArgs.Count < 2 Then Usage 2
sList = oArgs(0) : sPath = oArgs(1) : sFolder = ""
bJump = False : bQuiet = False
For i = 2 To oArgs.Count - 1
    sA = LCase(oArgs(i))
    Select Case sA
        Case "/jump"  : bJump  = True
        Case "/quiet" : bQuiet = True
        Case "/nolog" : g_LogEnabled = False
        Case Else
            If Left(oArgs(i),1) <> "/" Then
                If sFolder = "" Then sFolder = oArgs(i) Else Usage 2
            Else
                Usage 2
            End If
    End Select
Next
InitLog
LogMsg "INFO", "start: list=" & sList & " path=" & sPath

'--- INIT ---
Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
If Right(sPath, 1) <> "\" Then sPath = sPath & "\"

If sFolder = "" Then
    If bQuiet Then WScript.Quit 2
    sFolder = InputBox("Введите имя папки," & vbCrLf & "куда переместится выбранное:", _
                       "Перемещение выбранных элементов")
    If Trim(sFolder) = "" Then WScript.Quit 0
End If
Dim sTarg : sTarg = sPath & sFolder & "\"

'--- MAIN ---
On Error Resume Next
If Not oFso.FolderExists(sTarg) Then oFso.CreateFolder sTarg
If Err.Number <> 0 Then
    LogMsg "ERR", "mkdir fail: " & sTarg & " — " & Err.Description
    If Not bQuiet Then MsgBox "Не создать папку: " & Err.Description, vbCritical, "MoveFiles2NewFolders"
    WScript.Quit 1
End If
On Error GoTo 0

Dim oTs, sLine, iOk, iFail
Set oTs = oFso.OpenTextFile(sList, 1, False, -1)
iOk = 0 : iFail = 0
Do While Not oTs.AtEndOfStream
    sLine = Trim(oTs.ReadLine)
    If sLine <> "" Then
        On Error Resume Next
        If oFso.FolderExists(sLine) Then
            oFso.MoveFolder sLine, oFso.BuildPath(sTarg, oFso.GetFileName(sLine))
        ElseIf oFso.FileExists(sLine) Then
            oFso.MoveFile sLine, oFso.BuildPath(sTarg, oFso.GetFileName(sLine))
        Else
            Err.Number = 53
        End If
        If Err.Number <> 0 Then
            LogMsg "ERR", "move fail: " & sLine & " — " & Err.Description
            iFail = iFail + 1 : Err.Clear
        Else
            LogMsg "INFO", "move->" & sTarg
            iOk = iOk + 1
        End If
        On Error GoTo 0
    End If
Loop
oTs.Close
LogMsg "INFO", "moved=" & iOk & " failed=" & iFail

'--- EXIT ---
If bJump Then
    Dim oSh : Set oSh = CreateObject("WScript.Shell")
    Dim sSide
    If LCase(sPath) = LCase(oSh.CurrentDirectory & "\") Then sSide = "L" Else sSide = "R"
    oSh.Run """" & oSh.ExpandEnvironmentStrings("%COMMANDER_EXE%") & """ /O /S /" & sSide & "=""" & sTarg & """", 0, True
    If sSide = "R" Then oSh.SendKeys "{TAB}"
    LogMsg "INFO", "jump: " & sTarg
End If
If iFail > 0 Then WScript.Quit 1
WScript.Quit 0

'--- HELPERS ---
Sub Usage(code)
    WScript.Echo "MoveFiles2NewFolders.vbs <ListFile> <Path> [Folder] [/Jump] [/Quiet] [/NoLog]"
    WScript.Quit code
End Sub

'--- LOG ---
Sub InitLog()
    If Not g_LogEnabled Then Exit Sub
    Dim oFs, sDir, sCmdPath
    Set oFs = CreateObject("Scripting.FileSystemObject")
    sCmdPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%COMMANDER_PATH%")
    If sCmdPath <> "" And sCmdPath <> "%COMMANDER_PATH%" Then
        sDir = sCmdPath & "\PROFILE\LOG"
    Else
        sDir = oFs.GetSpecialFolder(2) & "\TCScripts"
    End If
    If Not oFs.FolderExists(sDir) Then
        On Error Resume Next
        oFs.CreateFolder sDir
        On Error GoTo 0
    End If
    If oFs.FolderExists(sDir) Then
        g_LogFile = oFs.BuildPath(sDir, SCRIPT_NAME & "_" & Year(Now) & Right("0" & Month(Now),2) & Right("0" & Day(Now),2) & ".log")
    End If
End Sub

Sub LogMsg(sLevel, sText)
    If Not g_LogEnabled Or g_LogFile = "" Then Exit Sub
    On Error Resume Next
    Dim oFs, oTs
    Set oFs = CreateObject("Scripting.FileSystemObject")
    Set oTs = oFs.OpenTextFile(g_LogFile, 8, True, -1)
    oTs.WriteLine FormatDateTime(Now, 0) & " [" & sLevel & "] " & sText
    oTs.Close
    On Error GoTo 0
End Sub