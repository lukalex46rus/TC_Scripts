'===============================================================================
' Script:      NamesToClip.vbs
' Version:     4.0
' Purpose:     Скопировать имена из списка (%L) в буфер обмена
' Description: Читает список, извлекает имена (опц. без расширений, опц. только
'              имена), при необходимости оборачивает в кавычки. Результат — в
'              буфер. Опционально открывает Notepad++.
'
' SYNTAX:      cscript //nologo NamesToClip.vbs <ListFile> [ключи]
' PARAMETERS:  <ListFile> — обычно %L
'
' KEYS:
'   /NoExt       - убрать расширения
'   /NameOnly    - только имя файла, без пути
'   /Quote       - обернуть строки с пробелами/& в кавычки
'   /Npp         - открыть Notepad++ на temp-файле
'   /KeepTemp    - не удалять temp-файл
'   /NoLog       - отключить лог
'
' RETURN CODES: 0 успех, 1 ошибка, 2 аргументы
' DEPENDENCIES: FSO, WScript.Shell, VBScript.RegExp
'===============================================================================
Option Explicit

Const SCRIPT_NAME = "NamesToClip"
Const NPP_PATH_REL = "\PROFILE\APP\Notepad++\notepad++.exe"

Dim g_LogEnabled : g_LogEnabled = True
Dim g_LogFile    : g_LogFile = ""
Dim g_StartTime  : g_StartTime = Now

'--- ARGS ---
Dim oArgs, sList, bNoExt, bNameOnly, bQuote, bNpp, bKeep, i, sA
Set oArgs = WScript.Arguments
If oArgs.Count = 0 Then Usage 2
sList = oArgs(0)
bNoExt = False : bNameOnly = False : bQuote = False : bNpp = False : bKeep = False
For i = 1 To oArgs.Count - 1
    sA = LCase(oArgs(i))
    Select Case sA
        Case "/noext"    : bNoExt    = True
        Case "/nameonly" : bNameOnly = True
        Case "/quote"    : bQuote    = True
        Case "/npp"      : bNpp      = True
        Case "/keeptemp" : bKeep     = True
        Case "/nolog"    : g_LogEnabled = False
        Case Else        : Usage 2
    End Select
Next
InitLog
LogMsg "INFO", "start: " & sList

'--- INIT ---
Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
If Not oFso.FileExists(sList) Then
    LogMsg "ERR", "list not found: " & sList
    WScript.Quit 1
End If

Dim oRe : Set oRe = New RegExp
oRe.Global = True : oRe.MultiLine = True : oRe.IgnoreCase = True

Dim sText
sText = oFso.OpenTextFile(sList, 1, False, -1).ReadAll
sText = Replace(sText, vbCrLf, vbLf)
sText = Replace(sText, vbCr, vbLf)
Do While Right(sText, 1) = vbLf : sText = Left(sText, Len(sText)-1) : Loop

'--- MAIN ---
If bNoExt Then
    If bNameOnly Then
        oRe.Pattern = "\.[^.\n\\]+(?=\\|$)"
    Else
        oRe.Pattern = "\.[^.\n]+$"
    End If
    sText = oRe.Replace(sText, "")
    LogMsg "INFO", "noext applied (nameonly=" & bNameOnly & ")"
End If

If bQuote Then
    oRe.Pattern = "^([^\r\n]*[ \t&][^\r\n]*)$"
    sText = oRe.Replace(sText, """$1""")
    LogMsg "INFO", "quote applied"
End If

Dim sTmp
sTmp = TempPath("names_" & Replace(CStr(Timer), ",", "") & ".txt")
oFso.CreateTextFile(sTmp, True, True).Write sText
LogMsg "INFO", "temp written: " & sTmp & " (" & Len(sText) & " chars)"

'--- CLIPBOARD ---
Dim oSh : Set oSh = CreateObject("WScript.Shell")
oSh.Run "%ComSpec% /c clip<""" & sTmp & """", 0, True
LogMsg "INFO", "clipboard updated"

'--- NPP ---
If bNpp Then
    Dim sNpp
    sNpp = oSh.ExpandEnvironmentStrings("%COMMANDER_PATH%") & NPP_PATH_REL
    If oFso.FileExists(sNpp) Then
        oSh.Run """" & sNpp & """ """ & sTmp & """", 1, False
        bKeep = True
        LogMsg "INFO", "npp launched: " & sNpp
    Else
        LogMsg "WARN", "npp not found: " & sNpp
    End If
End If

If Not bKeep Then oFso.DeleteFile sTmp, True
LogMsg "INFO", "done"
WScript.Quit 0

'--- HELPERS ---
Sub Usage(code)
    WScript.Echo "NamesToClip.vbs <ListFile> [/NoExt] [/NameOnly] [/Quote] [/Npp] [/KeepTemp] [/NoLog]"
    WScript.Quit code
End Sub

Function TempPath(sName)
    Dim sDir
    sDir = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%COMMANDER_PATH%") & "\PROFILE\TMP"
    If Not oFso.FolderExists(sDir) Then
        On Error Resume Next
        oFso.CreateFolder sDir
        On Error GoTo 0
    End If
    If oFso.FolderExists(sDir) Then
        TempPath = oFso.BuildPath(sDir, sName)
    Else
        TempPath = oFso.BuildPath(oFso.GetSpecialFolder(2), sName)
    End If
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