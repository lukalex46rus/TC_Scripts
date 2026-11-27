'===============================================================================
' Script:      TextClipboard2NewTXTFile.vbs
' Version:     4.0
' Purpose:     Текст из буфера обмена → новый .txt и открыть в TC
' Description: Имя-шаблон берётся из аргумента (обычно %N). При совпадении
'              создаётся «имя (N).txt». Кодировка выбирается автоматически.
'
' SYNTAX:      cscript //nologo TextClipboard2NewTXTFile.vbs <Template> [/NoLog]
' PARAMETERS:  <Template> — обычно %N
'
' KEYS:
'   /NoLog  - отключить лог
'
' RETURN CODES: 0 успех/пусто, 1 ошибка, 2 аргументы
' DEPENDENCIES: FSO, WScript.Shell, HTMLFile, VBScript.RegExp
'===============================================================================
Option Explicit

Const SCRIPT_NAME = "TextClipboard2NewTXTFile"
Dim g_LogEnabled : g_LogEnabled = True
Dim g_LogFile    : g_LogFile = ""
Dim g_StartTime  : g_StartTime = Now

'--- ARGS ---
Dim oArgs, sTemplate, i
Set oArgs = WScript.Arguments
If oArgs.Count = 0 Then WScript.Quit 2
sTemplate = oArgs(0)
For i = 1 To oArgs.Count - 1
    If LCase(oArgs(i)) = "/nolog" Then g_LogEnabled = False
Next
InitLog
LogMsg "INFO", "start: " & sTemplate

'--- INIT ---
Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
Dim oSh  : Set oSh  = CreateObject("WScript.Shell")

Dim sReg : sReg = "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1407"
Dim iOld, bRestore
On Error Resume Next
iOld = oSh.RegRead(sReg)
If Err.Number <> 0 Or iOld > 0 Then
    oSh.RegWrite sReg, 0, "REG_DWORD"
    bRestore = True
End If
Err.Clear
On Error GoTo 0

Dim sClip
On Error Resume Next
sClip = CreateObject("HTMLFile").parentWindow.clipboardData.GetData("text")
On Error GoTo 0

If bRestore Then oSh.RegWrite sReg, iOld, "REG_DWORD"

If IsNull(sClip) Or Trim(sClip) = "" Then
    LogMsg "INFO", "clipboard empty"
    WScript.Quit 0
End If

'--- MAIN ---
Dim oRe : Set oRe = New RegExp
oRe.Pattern = "[^ !-‚-›‘’–—\u0080-\u00FF\u0400-\u04FF\u20A0-\u20CF\u2100-\u214F\w]"
oRe.Global = True
Dim bEnc : bEnc = oRe.Test(sClip)

Dim sName, sExt, sBase, sFile, iN
sName = sTemplate
sExt  = oFso.GetExtensionName(sName)
sBase = oSh.CurrentDirectory & "\" & oFso.GetBaseName(sName)
sFile = sBase & "." & sExt
iN = 0
Do While oFso.FileExists(sFile) Or oFso.FolderExists(sFile)
    iN = iN + 1
    sFile = sBase & " (" & iN & ")." & sExt
Loop

oFso.CreateTextFile(sFile, False, bEnc).Write sClip
LogMsg "INFO", "written: " & sFile & " (unicode=" & bEnc & ")"

'--- EXIT ---
oSh.Exec """" & oSh.ExpandEnvironmentStrings("%COMMANDER_EXE%") & """ /A /O /S /L=""" & sFile & """"
LogMsg "INFO", "done"
WScript.Quit 0

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