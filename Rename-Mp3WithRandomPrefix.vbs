'===============================================================================
' Script:      Rename-Mp3WithRandomPrefix.vbs
' Version:     4.0
' Purpose:     Переименовать MP3, добавив случайный числовой префикс NNNNN-
' Description: Рекурсивно обходит каталог, для каждого .mp3 подбирает
'              свободное имя NNNNN-<оригинал>. Уже префиксованные пропускает
'              (если не /Force).
'
' SYNTAX:      cscript //nologo Rename-Mp3WithRandomPrefix.vbs <Path> [ключи]
' PARAMETERS:  <Path> — каталог с mp3
'
' KEYS:
'   /Recurse  - обрабатывать подпапки
'   /Force    - переименовывать уже префиксованные
'   /Min:N    - нижняя граница (по умолч. 10000)
'   /Max:N    - верхняя граница (по умолч. 32768)
'   /WhatIf   - сухой прогон
'   /NoLog    - отключить лог
'
' RETURN CODES: 0 успех, 1 ошибки, 2 аргументы, 3 путь не найден
' DEPENDENCIES: FSO, VBScript.RegExp
'===============================================================================
Option Explicit

Const SCRIPT_NAME = "Rename-Mp3WithRandomPrefix"
Const DEFAULT_MIN = 10000
Const DEFAULT_MAX = 32768
Const MAX_ATTEMPTS = 100
Const PREFIX_PATTERN = "^\d+-"

Dim g_LogEnabled : g_LogEnabled = True
Dim g_LogFile    : g_LogFile = ""
Dim g_StartTime  : g_StartTime = Now

'--- ARGS ---
Dim oArgs, sTarget, bRec, bForce, bWhatIf, iMin, iMax, i, sA
Set oArgs = WScript.Arguments
sTarget = "." : bRec = False : bForce = False : bWhatIf = False
iMin = DEFAULT_MIN : iMax = DEFAULT_MAX
For i = 0 To oArgs.Count - 1
    sA = LCase(oArgs(i))
    Select Case True
        Case sA = "/recurse" : bRec = True
        Case sA = "/force"   : bForce = True
        Case sA = "/whatif"  : bWhatIf = True
        Case sA = "/nolog"   : g_LogEnabled = False
        Case Left(sA,5) = "/min:" : iMin = CLng(Mid(oArgs(i), 6))
        Case Left(sA,5) = "/max:" : iMax = CLng(Mid(oArgs(i), 6))
        Case Left(oArgs(i),1) <> "/" : sTarget = oArgs(i)
        Case Else : Usage 2
    End Select
Next
If iMin > iMax Then
    WScript.Echo "Ошибка: Min > Max"
    WScript.Quit 2
End If
InitLog
LogMsg "INFO", "start: " & sTarget & " min=" & iMin & " max=" & iMax

'--- INIT ---
Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
If Not oFso.FolderExists(sTarget) Then
    LogMsg "ERR", "not found: " & sTarget
    WScript.Quit 3
End If
sTarget = oFso.GetAbsolutePathName(sTarget)

Dim oRe : Set oRe = New RegExp
oRe.Pattern = PREFIX_PATTERN : oRe.IgnoreCase = False

Randomize

Dim dFiles : Set dFiles = CreateObject("Scripting.Dictionary")
CollectMp3 sTarget, bRec, dFiles
If dFiles.Count = 0 Then
    LogMsg "INFO", "no mp3 in " & sTarget
    WScript.Echo "Нет .mp3 в " & sTarget
    WScript.Quit 0
End If

'--- MAIN ---
Dim aKeys, k, sPath, sName, sDir, sNew, iRen, iSkip, iFail
aKeys = dFiles.Keys
iRen = 0 : iSkip = 0 : iFail = 0

For k = 0 To UBound(aKeys)
    sPath = aKeys(k)
    sName = oFso.GetFileName(sPath)
    sDir  = oFso.GetParentFolderName(sPath)

    If oRe.Test(sName) And Not bForce Then
        iSkip = iSkip + 1
    Else
        sNew = GetRandomName(sName, sDir, iMin, iMax)
        If sNew = "" Then
            LogMsg "WARN", "no free name: " & sName
            WScript.Echo "Не подобрать имя: " & sName
            iFail = iFail + 1
        ElseIf bWhatIf Then
            WScript.Echo "[WhatIf] " & sName & "  →  " & sNew
            iRen = iRen + 1
        Else
            On Error Resume Next
            oFso.MoveFile sPath, oFso.BuildPath(sDir, sNew)
            If Err.Number <> 0 Then
                LogMsg "ERR", "rename fail: " & sName & " — " & Err.Description
                WScript.Echo "Ошибка: " & sName & " — " & Err.Description
                iFail = iFail + 1 : Err.Clear
            Else
                LogMsg "INFO", "ren: " & sName & " -> " & sNew
                WScript.Echo sName & "  →  " & sNew
                iRen = iRen + 1
            End If
            On Error GoTo 0
        End If
    End If
Next

'--- EXIT ---
WScript.Echo "Готово. Переименовано: " & iRen & ", пропущено: " & iSkip & ", ошибок: " & iFail
LogMsg "INFO", "renamed=" & iRen & " skipped=" & iSkip & " failed=" & iFail
If iFail > 0 Then WScript.Quit 1
WScript.Quit 0

'--- HELPERS ---
Sub Usage(code)
    WScript.Echo "Rename-Mp3WithRandomPrefix.vbs <Path> [/Recurse] [/Force] [/Min:N] [/Max:N] [/WhatIf] [/NoLog]"
    WScript.Quit code
End Sub

Sub CollectMp3(sPath, bR, dDict)
    Dim oF, oFile, oSub
    On Error Resume Next
    Set oF = oFso.GetFolder(sPath)
    If Err.Number <> 0 Then Err.Clear : On Error GoTo 0 : Exit Sub
    On Error GoTo 0
    For Each oFile In oF.Files
        If LCase(oFso.GetExtensionName(oFile.Name)) = "mp3" Then
            If Not dDict.Exists(oFile.Path) Then dDict.Add oFile.Path, True
        End If
    Next
    If bR Then
        For Each oSub In oF.SubFolders
            CollectMp3 oSub.Path, True, dDict
        Next
    End If
End Sub

Function GetRandomName(sOriginal, sDir, iLo, iHi)
    Dim n, i, sCand
    For i = 1 To MAX_ATTEMPTS
        n = Int((iHi - iLo + 1) * Rnd + iLo)
        sCand = CStr(n) & "-" & sOriginal
        If Not oFso.FileExists(oFso.BuildPath(sDir, sCand)) Then
            GetRandomName = sCand
            Exit Function
        End If
    Next
    GetRandomName = ""
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