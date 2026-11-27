'===============================================================================
' Script:      FileToOwnFolder.vbs
' Version:     4.0
' Purpose:     Переместить файл(ы) в папку с именем файла
' Description: Режим <Path> — один файл → папка рядом с ним.
'              Режим <ListFile> [BasePath] — каждый файл → своя папка.
'
' SYNTAX:
'   cscript //nologo FileToOwnFolder.vbs <Path>
'   cscript //nologo FileToOwnFolder.vbs <ListFile> [BasePath] [/Force] [/Quiet] [/NoLog]
'
' KEYS:
'   /Force  - перезаписывать без запроса
'   /Quiet  - без MsgBox
'   /NoLog  - отключить лог
'
' RETURN CODES: 0 успех, 1 ошибка, 2 аргументы, 3 не найдено
' DEPENDENCIES: FSO
'===============================================================================
Option Explicit

Const SCRIPT_NAME = "FileToOwnFolder"
Dim g_LogEnabled : g_LogEnabled = True
Dim g_LogFile    : g_LogFile = ""
Dim g_StartTime  : g_StartTime = Now

'--- ARGS ---
Dim oArgs, sFirst, sBase, bForce, bQuiet, i
Set oArgs = WScript.Arguments
If oArgs.Count = 0 Then Usage 2
sFirst = oArgs(0) : sBase = "" : bForce = False : bQuiet = False
For i = 1 To oArgs.Count - 1
    Select Case LCase(oArgs(i))
        Case "/force" : bForce = True
        Case "/quiet" : bQuiet = True
        Case "/nolog" : g_LogEnabled = False
        Case Else
            If Left(oArgs(i),1) <> "/" Then
                If sBase = "" Then sBase = oArgs(i) Else Usage 2
            Else
                Usage 2
            End If
    End Select
Next
InitLog
LogMsg "INFO", "start: " & sFirst

'--- INIT ---
Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
Dim bListMode : bListMode = oFso.FileExists(sFirst) And Not oFso.FolderExists(sFirst)

'--- MAIN ---
If bListMode Then
    Dim oTs, sLine
    Set oTs = oFso.OpenTextFile(sFirst, 1, False, -1)
    Do While Not oTs.AtEndOfStream
        sLine = Trim(oTs.ReadLine)
        If sLine <> "" Then ProcessOne sLine, sBase, bForce, bQuiet
    Loop
    oTs.Close
Else
    ProcessOne sFirst, sBase, bForce, bQuiet
End If
LogMsg "INFO", "done"
WScript.Quit 0

'--- HELPERS ---
Sub Usage(code)
    WScript.Echo "FileToOwnFolder.vbs <Path> | <ListFile> [BasePath] [/Force] [/Quiet] [/NoLog]"
    WScript.Quit code
End Sub

Sub ProcessOne(sPath, sBasePath, bF, bQ)
    If Not oFso.FileExists(sPath) Then
        LogMsg "WARN", "not found: " & sPath
        If Not bQ Then MsgBox "Не найден файл: " & sPath, vbExclamation, "FileToOwnFolder"
        Exit Sub
    End If

    Dim sName, sParent, sDestFolder, sDest
    sName   = oFso.GetFileName(sPath)
    sParent = oFso.GetParentFolderName(sPath)

    If sBasePath = "" Then
        sDestFolder = oFso.BuildPath(sParent, oFso.GetBaseName(sPath))
    Else
        If Not oFso.FolderExists(sBasePath) Then
            LogMsg "ERR", "no base: " & sBasePath
            If Not bQ Then MsgBox "Нет базовой папки: " & sBasePath, vbCritical, "FileToOwnFolder"
            Exit Sub
        End If
        sDestFolder = oFso.BuildPath(sBasePath, oFso.GetBaseName(sPath))
    End If

    On Error Resume Next
    If Not oFso.FolderExists(sDestFolder) Then oFso.CreateFolder sDestFolder
    If Err.Number <> 0 Then
        LogMsg "ERR", "mkdir fail: " & sDestFolder & " — " & Err.Description
        If Not bQ Then MsgBox "Не создать папку: " & Err.Description, vbCritical, "FileToOwnFolder"
        Err.Clear : On Error GoTo 0 : Exit Sub
    End If
    On Error GoTo 0

    sDest = oFso.BuildPath(sDestFolder, sName)

    If oFso.FileExists(sDest) Then
        Dim ans
        If bF Then
            ans = vbYes
        ElseIf bQ Then
            ans = vbNo
        Else
            ans = MsgBox("Файл уже есть:" & vbCrLf & sDest & vbCrLf & "Перезаписать?", _
                         vbYesNo + vbQuestion, "FileToOwnFolder")
        End If
        If ans <> vbYes Then
            LogMsg "INFO", "skip (exists): " & sDest
            Exit Sub
        End If
        On Error Resume Next
        oFso.DeleteFile sDest, True
        Err.Clear
        On Error GoTo 0
    End If

    On Error Resume Next
    oFso.MoveFile sPath, sDest
    If Err.Number <> 0 Then
        LogMsg "ERR", "move fail: " & sPath & " — " & Err.Description
        If Not bQ Then MsgBox "Не переместить: " & Err.Description, vbCritical, "FileToOwnFolder"
        Err.Clear
    Else
        LogMsg "INFO", "move->" & sDest
    End If
    On Error GoTo 0
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