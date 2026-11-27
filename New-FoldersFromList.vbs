'===============================================================================
' Script:      New-FoldersFromList.vbs
' Version:     4.0
' Purpose:     Массово создать папки из файла или буфера обмена
' Description: Список — по одному имени на строку. Вложенные пути (A\B\C)
'              создаются рекурсивно. Абсолютные пути и «..» отклоняются.
'              Для буфера сначала показывается план и запрос подтверждения.
'
' SYNTAX:
'   cscript //nologo New-FoldersFromList.vbs <ListFile> <BasePath> [/WhatIf] [/NoLog]
'   cscript //nologo New-FoldersFromList.vbs -ListClipboard <BasePath> [/WhatIf] [/NoLog]
'
' KEYS:
'   -ListClipboard - читать список из буфера
'   /WhatIf        - только показать план
'   /NoLog         - отключить лог
'
' RETURN CODES: 0 успех, 1 ошибки, 2 аргументы, 3 нет BasePath
' DEPENDENCIES: FSO, ADODB.Stream (опц.), HTMLFile (clipboard)
'===============================================================================
Option Explicit

Const SCRIPT_NAME = "New-FoldersFromList"
Dim g_LogEnabled : g_LogEnabled = True
Dim g_LogFile    : g_LogFile = ""
Dim g_StartTime  : g_StartTime = Now

'--- ARGS ---
Dim oArgs, sList, sBase, bWhatIf, bClip, i, sA
Set oArgs = WScript.Arguments
sList = "" : sBase = "" : bWhatIf = False : bClip = False
For i = 0 To oArgs.Count - 1
    sA = LCase(oArgs(i))
    Select Case sA
        Case "/whatif"        : bWhatIf = True
        Case "-listclipboard" : bClip   = True
        Case "/nolog"         : g_LogEnabled = False
        Case Else
            If sList = "" And Not bClip Then sList = oArgs(i)
            ElseIf sBase = "" Then sBase = oArgs(i)
    End Select
Next
If sBase = "" Then Usage 2
If Not bClip And sList = "" Then Usage 2
InitLog
LogMsg "INFO", "start: base=" & sBase & " clip=" & bClip

'--- INIT ---
Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
If Not oFso.FolderExists(sBase) Then
    LogMsg "ERR", "base not found: " & sBase
    WScript.Quit 3
End If
sBase = oFso.GetAbsolutePathName(sBase)
If Right(sBase, 1) <> "\" Then sBase = sBase & "\"

Dim aLines
If bClip Then
    aLines = ReadClipboard()
    If Not IsArray(aLines) Then
        LogMsg "WARN", "clipboard empty"
        WScript.Echo "Буфер пуст или недоступен."
        WScript.Quit 2
    End If
Else
    If Not oFso.FileExists(sList) Then
        LogMsg "ERR", "list not found: " & sList
        WScript.Quit 2
    End If
    aLines = ReadListFile(sList)
    If Not IsArray(aLines) Then WScript.Quit 2
End If

'--- MAIN ---
Dim bPreview : bPreview = bWhatIf Or bClip
Dim dStats : Set dStats = ProcessLines(aLines, bPreview)

If bWhatIf And Not bClip Then
    PrintReport dStats
    WScript.Quit 0
End If

If bClip Then
    PrintReport dStats
    If dStats("created") = 0 Then WScript.Quit 0
    Dim ans
    ans = MsgBox("Будет создано папок: " & dStats("created") & vbCrLf & _
                 "Пропущено: " & dStats("skipped") & vbCrLf & _
                 "Ошибок: " & dStats("failed") & vbCrLf & vbCrLf & "Выполнить?", _
                 vbQuestion + vbOKCancel + vbDefaultButton2, "Подтверждение")
    If ans <> vbOK Then WScript.Quit 0
    Set dStats = ProcessLines(aLines, False)
End If

'--- EXIT ---
PrintReport dStats
LogMsg "INFO", "created=" & dStats("created") & " skipped=" & dStats("skipped") & " failed=" & dStats("failed")
If dStats("failed") > 0 Then WScript.Quit 1
WScript.Quit 0

'--- HELPERS ---
Sub Usage(code)
    WScript.Echo "New-FoldersFromList.vbs <ListFile> <BasePath> [/WhatIf] [/NoLog]"
    WScript.Echo "New-FoldersFromList.vbs -ListClipboard <BasePath> [/WhatIf] [/NoLog]"
    WScript.Quit code
End Sub

Function ProcessLines(aArr, bPreview)
    Dim dRes, i, sRaw, sName, sFull, sBaseSep
    Set dRes = CreateObject("Scripting.Dictionary")
    dRes.Add "created", 0 : dRes.Add "skipped", 0 : dRes.Add "failed", 0
    sBaseSep = sBase

    For i = 0 To UBound(aArr)
        sRaw = aArr(i)
        If IsNull(sRaw) Then GoTo NextLine
        sName = Trim(CStr(sRaw))
        If sName = "" Then GoTo NextLine
        Do While Right(sName,1) = "\" Or Right(sName,1) = "/"
            sName = Left(sName, Len(sName)-1)
        Loop
        If sName = "" Then GoTo NextLine

        If IsPathRooted(sName) Then
            dRes("skipped") = dRes("skipped") + 1 : GoTo NextLine
        End If
        If InStr(sName, "..") > 0 Then
            dRes("skipped") = dRes("skipped") + 1 : GoTo NextLine
        End If

        On Error Resume Next
        sFull = oFso.GetAbsolutePathName(oFso.BuildPath(sBase, sName))
        If Err.Number <> 0 Then
            dRes("failed") = dRes("failed") + 1 : Err.Clear : On Error GoTo 0 : GoTo NextLine
        End If
        On Error GoTo 0

        If LCase(Left(sFull, Len(sBaseSep))) <> LCase(sBaseSep) Then
            dRes("skipped") = dRes("skipped") + 1 : GoTo NextLine
        End If

        If oFso.FolderExists(sFull) Then
            dRes("skipped") = dRes("skipped") + 1 : GoTo NextLine
        End If

        If bPreview Then
            WScript.Echo "[WhatIf] " & sFull
            dRes("created") = dRes("created") + 1
        Else
            On Error Resume Next
            MakeDir sFull
            If Err.Number <> 0 Then
                LogMsg "ERR", "mkdir fail: " & sFull & " — " & Err.Description
                WScript.Echo "Ошибка: " & sFull & " — " & Err.Description
                dRes("failed") = dRes("failed") + 1 : Err.Clear
            Else
                LogMsg "INFO", "mkdir: " & sFull
                WScript.Echo "Создана: " & sFull
                dRes("created") = dRes("created") + 1
            End If
            On Error GoTo 0
        End If
NextLine:
    Next
    Set ProcessLines = dRes
End Function

Sub PrintReport(dS)
    WScript.Echo "──────────────────────────────────────"
    WScript.Echo "Создано: " & dS("created") & _
                 ", пропущено: " & dS("skipped") & _
                 ", ошибок: " & dS("failed")
    WScript.Echo "──────────────────────────────────────"
End Sub

Function ReadListFile(sPath)
    Dim oStream, sContent
    On Error Resume Next
    Set oStream = CreateObject("ADODB.Stream")
    If Err.Number = 0 Then
        oStream.Type = 2 : oStream.Charset = "utf-8" : oStream.Open
        oStream.LoadFromFile sPath
        If Err.Number = 0 Then
            sContent = oStream.ReadText(-1)
            oStream.Close
            On Error GoTo 0
            If Left(sContent,1) = ChrW(&HFEFF) Then sContent = Mid(sContent,2)
            ReadListFile = Normalize(sContent)
            Exit Function
        End If
        Err.Clear
    End If
    On Error GoTo 0

    On Error Resume Next
    Dim oTs
    Set oTs = oFso.OpenTextFile(sPath, 1, False, 0)
    If Err.Number <> 0 Then ReadListFile = Empty : Exit Function
    sContent = oTs.ReadAll
    oTs.Close
    On Error GoTo 0
    ReadListFile = Normalize(sContent)
End Function

Function ReadClipboard()
    Dim oHtml, sClip
    On Error Resume Next
    Set oHtml = CreateObject("HTMLFile")
    If Err.Number <> 0 Then ReadClipboard = Empty : Exit Function
    oHtml.Write "<html><body></body></html>"
    sClip = oHtml.parentWindow.clipboardData.GetData("text")
    If Err.Number <> 0 Or IsNull(sClip) Or IsEmpty(sClip) Then
        ReadClipboard = Empty
    Else
        ReadClipboard = Normalize(CStr(sClip))
    End If
    Set oHtml = Nothing
    On Error GoTo 0
End Function

Function Normalize(sText)
    Dim s
    s = Replace(sText, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    Normalize = Split(s, vbLf)
End Function

Function IsPathRooted(sP)
    IsPathRooted = False
    If Len(sP) >= 2 And Mid(sP, 2, 1) = ":" Then IsPathRooted = True
    If Len(sP) >= 2 And Left(sP, 2) = "\\" Then IsPathRooted = True
End Function

Sub MakeDir(sFull)
    Dim sParent
    sParent = oFso.GetParentFolderName(sFull)
    If sParent <> "" And Not oFso.FolderExists(sParent) Then MakeDir sParent
    If Not oFso.FolderExists(sFull) Then oFso.CreateFolder sFull
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