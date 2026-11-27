'===============================================================================
' Script:      MoveUp.vbs
' Version:     4.0
' Purpose:     Поднять файлы/папки на один уровень выше
' Description: Принимает одиночный путь либо список (%L). Для каждого элемента
'              находит «деда» и переносит туда. При коллизии имён создаёт копию
'              с суффиксом _Copy (файл/папка).
'
' SYNTAX:
'   cscript //nologo MoveUp.vbs <Path>
'   cscript //nologo MoveUp.vbs <ListFile> [/Jump] [/Quiet] [/NoLog]
'
' PARAMETERS:
'   <Path>      - один файл/папка
'   <ListFile>  - список путей (обычно %L)
'
' KEYS:
'   /Jump   - после операции перейти в целевую папку
'   /Quiet  - без MsgBox
'   /NoLog  - отключить лог (по умолчанию лог включён)
'
' RETURN CODES: 0 успех, 1 ошибка, 2 аргументы, 3 не найдено
' DEPENDENCIES: FSO, WScript.Shell
'===============================================================================
Option Explicit

Const SCRIPT_NAME = "MoveUp"

'--- LOG INIT ---
Dim g_LogEnabled : g_LogEnabled = True
Dim g_LogFile    : g_LogFile = ""
Dim g_StartTime  : g_StartTime = Now

'--- ARGS ---
Dim oArgs, sFirst, bJump, bQuiet, i
Set oArgs = WScript.Arguments
If oArgs.Count = 0 Then Usage 2
sFirst = oArgs(0) : bJump = False : bQuiet = False
For i = 1 To oArgs.Count - 1
    Select Case LCase(oArgs(i))
        Case "/jump"  : bJump  = True
        Case "/quiet" : bQuiet = True
        Case "/nolog" : g_LogEnabled = False
        Case Else     : Usage 2
    End Select
Next

InitLog
LogMsg "INFO", "start: " & sFirst

'--- INIT ---
Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
Dim sLastTarget : sLastTarget = ""
Dim bListMode : bListMode = oFso.FileExists(sFirst) And Not oFso.FolderExists(sFirst)

'--- MAIN ---
If bListMode Then
    Dim oTs, sLine
    Set oTs = oFso.OpenTextFile(sFirst, 1, False, -1)
    Do While Not oTs.AtEndOfStream
        sLine = Trim(oTs.ReadLine)
        If sLine <> "" Then sLastTarget = MoveOne(sLine, bQuiet, sLastTarget)
    Loop
    oTs.Close
Else
    sLastTarget = MoveOne(sFirst, bQuiet, "")
End If

'--- EXIT ---
If bJump And sLastTarget <> "" Then
    Dim oSh : Set oSh = CreateObject("WScript.Shell")
    oSh.Run """" & oSh.ExpandEnvironmentStrings("%COMMANDER_EXE%") & """ /O /S /L=""" & sLastTarget & """", 0, False
    LogMsg "INFO", "jump: " & sLastTarget
End If
LogMsg "INFO", "done"
WScript.Quit 0

'--- HELPERS ---
Sub Usage(code)
    WScript.Echo "MoveUp.vbs <Path> | <ListFile> [/Jump] [/Quiet] [/NoLog]"
    WScript.Quit code
End Sub

Function MoveOne(sItem, bQ, sPrev)
    Dim bFolder, sName, sParent, sGrand, sDest
    MoveOne = sPrev

    If oFso.FolderExists(sItem) Then
        bFolder = True
    ElseIf oFso.FileExists(sItem) Then
        bFolder = False
    Else
        LogMsg "WARN", "not found: " & sItem
        If Not bQ Then MsgBox "Не найдено: " & sItem, vbExclamation, "MoveUp"
        Exit Function
    End If

    sName   = oFso.GetFileName(sItem)
    sParent = oFso.GetParentFolderName(sItem)
    If sParent = "" Then Exit Function
    sGrand  = oFso.GetParentFolderName(sParent)
    If sGrand  = "" Then
        LogMsg "WARN", "no grand parent: " & sItem
        Exit Function
    End If

    sDest = oFso.BuildPath(sGrand, sName)
    MoveOne = sGrand
    If LCase(sDest) = LCase(sItem) Then Exit Function

    On Error Resume Next
    If oFso.FolderExists(sDest) Or oFso.FileExists(sDest) Then
        Dim sCopy
        If bFolder Then
            sCopy = oFso.BuildPath(sGrand, sName & "_Copy")
            oFso.CopyFolder sItem, sCopy
            If Err.Number = 0 Then oFso.DeleteFolder sItem, True
        Else
            sCopy = oFso.BuildPath(sGrand, oFso.GetBaseName(sName) & "_Copy." & oFso.GetExtensionName(sName))
            oFso.CopyFile sItem, sCopy
            If Err.Number = 0 Then oFso.DeleteFile sItem, True
        End If
        If Err.Number = 0 Then LogMsg "INFO", "copy->" & sCopy & " (collision)" Else LogMsg "ERR", "copy fail: " & Err.Description
    Else
        If bFolder Then oFso.MoveFolder sItem, sDest Else oFso.MoveFile sItem, sDest
        If Err.Number = 0 Then LogMsg "INFO", "move->" & sDest Else LogMsg "ERR", "move fail: " & Err.Description
    End If
    If Err.Number <> 0 Then
        If Not bQ Then MsgBox "Ошибка: " & Err.Description, vbCritical, "MoveUp"
        Err.Clear
    End If
    On Error GoTo 0
End Function

'--- LOG ---
Sub InitLog()
    If Not g_LogEnabled Then Exit Sub
    Dim oFso, sDir, sCmdPath
    Set oFso = CreateObject("Scripting.FileSystemObject")
    sCmdPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%COMMANDER_PATH%")
    If sCmdPath <> "" And sCmdPath <> "%COMMANDER_PATH%" Then
        sDir = sCmdPath & "\PROFILE\LOG"
    Else
        sDir = oFso.GetSpecialFolder(2) & "\TCScripts"
    End If
    If Not oFso.FolderExists(sDir) Then
        On Error Resume Next
        oFso.CreateFolder sDir
        On Error GoTo 0
    End If
    If oFso.FolderExists(sDir) Then
        g_LogFile = oFso.BuildPath(sDir, SCRIPT_NAME & "_" & Year(Now) & Right("0" & Month(Now),2) & Right("0" & Day(Now),2) & ".log")
    End If
End Sub

Sub LogMsg(sLevel, sText)
    If Not g_LogEnabled Or g_LogFile = "" Then Exit Sub
    On Error Resume Next
    Dim oFso, oTs
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Set oTs = oFso.OpenTextFile(g_LogFile, 8, True, -1)
    oTs.WriteLine FormatDateTime(Now, 0) & " [" & sLevel & "] " & sText
    oTs.Close
    On Error GoTo 0
End Sub