'===============================================================================
' Script:      FileList.vbs
' Version:     4.0
' Purpose:     Сформировать список файлов/папок в целевой файл
' Description: Источник — папка или список путей. Для папки — рекурсивно или
'              нет. Пути можно делать относительными к источнику. /New —
'              создать новый файл с суффиксом, если целевой занят.
'
' SYNTAX:
'   cscript //nologo FileList.vbs <Source> <TargetList> [/Recurse] [/Relative] [/New] [/NoLog]
'
' PARAMETERS:
'   <Source>      - папка или список путей (%L)
'   <TargetList>  - файл-результат (относительно %P, если не абсолютный)
'
' KEYS:
'   /Recurse  - обходить подпапки
'   /Relative - относительные пути (относительно источника)
'   /New      - создать файл с суффиксом _NN при коллизии
'   /NoLog    - отключить лог
'
' RETURN CODES: 0 успех, 2 аргументы, 3 источник не найден
' DEPENDENCIES: FSO, WScript.Shell
'===============================================================================
Option Explicit

Const SCRIPT_NAME = "FileList"
Dim g_LogEnabled : g_LogEnabled = True
Dim g_LogFile    : g_LogFile = ""
Dim g_StartTime  : g_StartTime = Now

'--- ARGS ---
Dim oArgs, sSrc, sTarget, bRec, bRel, bNew, i, sA
Set oArgs = WScript.Arguments
If oArgs.Count < 2 Then Usage 2
sSrc = oArgs(0) : sTarget = oArgs(1)
bRec = False : bRel = False : bNew = False
For i = 2 To oArgs.Count - 1
    sA = LCase(oArgs(i))
    Select Case sA
        Case "/recurse"  : bRec = True
        Case "/relative" : bRel = True
        Case "/new"      : bNew = True
        Case "/nolog"    : g_LogEnabled = False
        Case Else        : Usage 2
    End Select
Next
InitLog
LogMsg "INFO", "start: src=" & sSrc & " target=" & sTarget

'--- INIT ---
Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
If Not oFso.FolderExists(sSrc) And Not oFso.FileExists(sSrc) Then
    LogMsg "ERR", "src not found: " & sSrc
    WScript.Quit 3
End If

Dim bFolderMode : bFolderMode = oFso.FolderExists(sSrc)
Dim sCur : sCur = CreateObject("WScript.Shell").CurrentDirectory
If bFolderMode Then sCur = sSrc
If Right(sCur, 1) <> "\" Then sCur = sCur & "\"
Dim iLen : iLen = Len(sCur)

If InStr(sTarget, ":") = 0 Then sTarget = sCur & sTarget
If bNew Then sTarget = NextName(sTarget)

'--- MAIN ---
Dim sOut : sOut = ""
If bFolderMode Then
    sOut = WalkFolder(sSrc, bRec, bRel)
Else
    Dim oTs, sLine
    Set oTs = oFso.OpenTextFile(sSrc, 1, False, -1)
    Do While Not oTs.AtEndOfStream
        sLine = Trim(oTs.ReadLine)
        If sLine <> "" Then sOut = sOut & Cut(sLine, bRel) & vbNewLine
    Loop
    oTs.Close
End If

oFso.CreateTextFile(sTarget, True, True).Write sOut
LogMsg "INFO", "written: " & sTarget & " (" & Len(sOut) & " chars)"
LogMsg "INFO", "done"
WScript.Quit 0

'--- HELPERS ---
Sub Usage(code)
    WScript.Echo "FileList.vbs <Source> <TargetList> [/Recurse] [/Relative] [/New] [/NoLog]"
    WScript.Quit code
End Sub

Function WalkFolder(sFolder, bR, bRel)
    Dim sRes, oFolder, oFile, oSub
    sRes = ""
    Set oFolder = oFso.GetFolder(sFolder)
    If bR Then
        For Each oFile In oFolder.Files
            sRes = sRes & Cut(oFile.Path, bRel) & vbNewLine
        Next
        For Each oSub In oFolder.SubFolders
            sRes = sRes & Cut(oSub.Path, bRel) & vbNewLine
            sRes = sRes & WalkFolder(oSub.Path, True, bRel)
        Next
    Else
        sRes = sRes & Cut(sFolder, bRel) & vbNewLine
    End If
    WalkFolder = sRes
End Function

Function Cut(sPath, bRelative)
    Dim sAbs : sAbs = oFso.GetAbsolutePathName(sPath)
    If bRelative Then
        If LCase(Left(sAbs, iLen)) = LCase(sCur) Then
            Cut = Mid(sAbs, iLen + 1)
        Else
            Cut = sAbs
        End If
    Else
        Cut = sAbs
    End If
End Function

Function NextName(sPath)
    Const lQ = 2
    Dim sDir, sBase, sExt, sTry, iN, sAdd
    sDir  = oFso.GetParentFolderName(sPath)
    If sDir <> "" Then sDir = sDir & "\"
    sBase = oFso.GetBaseName(sPath)
    sExt  = oFso.GetExtensionName(sPath)
    NextName = sPath
    If Not oFso.FileExists(NextName) Then Exit Function
    Do
        iN = iN + 1
        If iN < 10 ^ lQ Then sAdd = Right(String(lQ, "0") & iN, lQ) Else sAdd = CStr(iN)
        NextName = sDir & sBase & "_" & sAdd & "." & sExt
    Loop While oFso.FileExists(NextName)
End Function

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