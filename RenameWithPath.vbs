'===============================================================================
' Script:      RenameWithPath.vbs
' Version:     4.0
' Purpose:     Переименовать файл/папку, при необходимости создав подпапки
' Description: В InputBox можно ввести «A/B/C/NewName.txt» — будут созданы
'              все недостающие папки, объект перемещён под новое имя.
'
' SYNTAX:      cscript //nologo RenameWithPath.vbs <Path> [/NoLog]
' PARAMETERS:  <Path> — файл или папка
'
' KEYS:
'   /NoLog  - отключить лог
'
' RETURN CODES: 0 успех, 1 ошибка, 2 аргументы, 3 не найдено
' DEPENDENCIES: FSO
'===============================================================================
Option Explicit

Const SCRIPT_NAME = "RenameWithPath"
Dim g_LogEnabled : g_LogEnabled = True
Dim g_LogFile    : g_LogFile = ""
Dim g_StartTime  : g_StartTime = Now

'--- ARGS ---
Dim oArgs, sPath, i
Set oArgs = WScript.Arguments
If oArgs.Count = 0 Then Usage 2
sPath = oArgs(0)
For i = 1 To oArgs.Count - 1
    If LCase(oArgs(i)) = "/nolog" Then g_LogEnabled = False Else Usage 2
Next
InitLog
LogMsg "INFO", "start: " & sPath

'--- INIT ---
Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
Dim bFolder
If oFso.FolderExists(sPath) Then
    bFolder = True
ElseIf oFso.FileExists(sPath) Then
    bFolder = False
Else
    LogMsg "ERR", "not found: " & sPath
    WScript.Quit 3
End If

Dim sOldName, sParent
If bFolder Then
    sOldName = oFso.GetFolder(sPath).Name
Else
    sOldName = oFso.GetFileName(sPath)
End If
sParent = oFso.GetParentFolderName(sPath)

'--- MAIN ---
Dim sInput
sInput = InputBox("Введите новое имя (можно с / или \ для создания папок):" & vbCrLf & vbCrLf & _
                  "Пример: F1/FF2/FF3/NewName.txt", "Переименование с созданием папок", sOldName)
If Trim(sInput) = "" Then WScript.Quit 0
If StrComp(sInput, sOldName, vbTextCompare) = 0 Then WScript.Quit 0

Dim sNorm
sNorm = Replace(sInput, "/", "\")
Do While InStr(sNorm, "\\") > 0 : sNorm = Replace(sNorm, "\\", "\") : Loop
Do While Left(sNorm, 1) = "\" : sNorm = Mid(sNorm, 2) : Loop
Do While Right(sNorm, 1) = "\" : sNorm = Left(sNorm, Len(sNorm)-1) : Loop
If sNorm = "" Then WScript.Quit 1

Dim aParts, sTarget
aParts = Split(sNorm, "\")
Dim iFolders : iFolders = UBound(aParts)
sTarget = aParts(UBound(aParts))
If sTarget = "" Then WScript.Quit 1

Dim aBad : aBad = Array(":", "*", "?", """", "<", ">", "|")
Dim j
For j = 0 To UBound(aBad)
    If InStr(sTarget, aBad(j)) > 0 Then WScript.Quit 1
Next
For i = 0 To iFolders - 1
    If aParts(i) = "" Then WScript.Quit 1
    For j = 0 To UBound(aBad)
        If InStr(aParts(i), aBad(j)) > 0 Then WScript.Quit 1
    Next
Next

Dim sCur : sCur = sParent
For i = 0 To iFolders - 1
    sCur = oFso.BuildPath(sCur, aParts(i))
    If Not oFso.FolderExists(sCur) Then
        On Error Resume Next
        oFso.CreateFolder sCur
        If Err.Number <> 0 Then WScript.Quit 1
        On Error GoTo 0
    End If
Next

Dim sDest : sDest = oFso.BuildPath(sCur, sTarget)
If LCase(sDest) = LCase(sPath) Then WScript.Quit 0

If bFolder Then
    If LCase(Left(sDest, Len(sPath))) = LCase(sPath) Then WScript.Quit 1
End If

'--- EXIT ---
On Error Resume Next
If bFolder Then
    oFso.MoveFolder sPath, sDest
Else
    If oFso.FileExists(sDest) Then
        If MsgBox("Файл уже существует:" & vbCrLf & sDest & vbCrLf & "Перезаписать?", _
                  vbYesNo + vbQuestion, "Подтверждение") <> vbYes Then WScript.Quit 0
        oFso.DeleteFile sDest, True
    End If
    oFso.MoveFile sPath, sDest
End If
If Err.Number <> 0 Then
    LogMsg "ERR", "move fail: " & Err.Description
    WScript.Quit 1
End If
LogMsg "INFO", "ren: " & sPath & " -> " & sDest
WScript.Quit 0

'--- HELPERS ---
Sub Usage(code)
    WScript.Echo "RenameWithPath.vbs <Path> [/NoLog]"
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