'===============================================================================
' Script:      MirrorFolderTree.vbs
' Version:     4.0
' Purpose:     Зеркалировать дерево папок на неактивную панель
' Description: Читает %L (только папки), создаёт одноимённые папки в %T и
'              рекурсивно повторяет поддерево. Существующие папки пропускает.
'
' SYNTAX:      cscript //nologo MirrorFolderTree.vbs <ListFile> <DestRoot> [/Quiet] [/NoLog]
' PARAMETERS:  <ListFile> = %L, <DestRoot> = %T
'
' KEYS:
'   /Quiet  - без итогового MsgBox
'   /NoLog  - отключить лог
'
' RETURN CODES: 0 успех, 1 ошибка, 2 аргументы, 3 путь не найден
' DEPENDENCIES: FSO
'===============================================================================
Option Explicit

Const SCRIPT_NAME = "MirrorFolderTree"
Dim g_LogEnabled : g_LogEnabled = True
Dim g_LogFile    : g_LogFile = ""
Dim g_StartTime  : g_StartTime = Now

'--- ARGS ---
Dim oArgs, sList, sDst, bQuiet, i
Set oArgs = WScript.Arguments
If oArgs.Count < 2 Then Usage 2
sList = oArgs(0) : sDst = oArgs(1) : bQuiet = False
For i = 2 To oArgs.Count - 1
    Select Case LCase(oArgs(i))
        Case "/quiet" : bQuiet = True
        Case "/nolog" : g_LogEnabled = False
        Case Else     : Usage 2
    End Select
Next
InitLog
LogMsg "INFO", "start: " & sList & " -> " & sDst

'--- INIT ---
Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
If Not oFso.FileExists(sList) Then
    LogMsg "ERR", "list not found: " & sList
    WScript.Quit 3
End If
If Not oFso.FolderExists(sDst) Then
    LogMsg "ERR", "dst not found: " & sDst
    WScript.Quit 3
End If

Dim iFolders : iFolders = 0
Dim iSubs    : iSubs    = 0

'--- MAIN ---
Dim oTs, sLine
Set oTs = oFso.OpenTextFile(sList, 1, False, -1)
Do While Not oTs.AtEndOfStream
    sLine = Trim(oTs.ReadLine)
    If Len(sLine) >= 2 Then
        If Left(sLine, 1) = """" And Right(sLine, 1) = """" Then
            sLine = Mid(sLine, 2, Len(sLine) - 2)
        End If
    End If
    If sLine <> "" And oFso.FolderExists(sLine) Then
        iFolders = iFolders + 1
        MirrorOne sLine
    End If
Loop
oTs.Close
LogMsg "INFO", "folders=" & iFolders & " subcreated=" & iSubs

'--- EXIT ---
If Not bQuiet Then
    MsgBox "Готово." & vbCrLf & _
           "Обработано папок: " & iFolders & vbCrLf & _
           "Создано подпапок: " & iSubs, vbInformation, "MirrorFolderTree"
End If
WScript.Quit 0

'--- HELPERS ---
Sub Usage(code)
    WScript.Echo "MirrorFolderTree.vbs <ListFile> <DestRoot> [/Quiet] [/NoLog]"
    WScript.Quit code
End Sub

Sub MirrorOne(sSrc)
    Dim oSrc, sPath
    Set oSrc = oFso.GetFolder(sSrc)
    sPath = oFso.BuildPath(sDst, oSrc.Name)
    If Not oFso.FolderExists(sPath) Then
        On Error Resume Next
        oFso.CreateFolder sPath
        If Err.Number = 0 Then LogMsg "INFO", "mkdir: " & sPath Else LogMsg "ERR", "mkdir fail: " & sPath
        Err.Clear
        On Error GoTo 0
    End If
    CopyTree oSrc, sPath
End Sub

Sub CopyTree(oSrc, sDstFld)
    Dim oSub, sNew
    For Each oSub In oSrc.SubFolders
        sNew = oFso.BuildPath(sDstFld, oSub.Name)
        If Not oFso.FolderExists(sNew) Then
            On Error Resume Next
            oFso.CreateFolder sNew
            If Err.Number = 0 Then
                iSubs = iSubs + 1
                LogMsg "INFO", "mkdir: " & sNew
            Else
                LogMsg "ERR", "mkdir fail: " & sNew
            End If
            Err.Clear
            On Error GoTo 0
        Else
            iSubs = iSubs + 1
        End If
        CopyTree oSub, sNew
    Next
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