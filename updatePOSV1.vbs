Option Explicit

' Global Configuration
Const SourceDir = "http://192.168.1.99/"
Const DestDskTopDir = "%USERPROFILE%\Desktop"
Const DestPrgmDir = "%PROGRAMFILES%\BarMade"
Dim SourceFiles : SourceFiles = Array( "BarMade.exe", "BarMade.exe", "BarMadeDispatcher.exe" )
Dim DestDirs : DestDirs = Array( DestPrgmDir, DestDskTopDir, DestPrgmDir )
Dim ScriptDir : ScriptDir = GetScriptDir()
Dim CurrentDir : CurrentDir = GetCurrentDir()

WScript.Quit ExecuteUpdateProcess()

'****************************************************************************************

Function ExecuteUpdateProcess()
  On Error Resume Next
  Dim ErrMsg, Operation
  ExecuteUpdateProcess = 0
  ErrMsg = ArgumentsOK(Operation)
  If ErrMsg <> "" then
    Call ShowUsage(ErrMsg)
  Else
    if operation = "CHECKFORUPDATE" then
      Call CheckForUpdate()
    elseif operation = "UPDATEPOS" then
      Call UpdatePOS()
    end if
  End If
  if Err.Number <> 0 Then
    ExecuteUpdateProcess = 1
    WScript.Echo "Error: " & Err.Number & " " & Err.Description
  end if
End Function

'***************************************************************************************

Sub ShowUsage(ByVal ErrMsg)
  Dim output : output = ""
  If ErrMsg <> "" Then
    output = output & "Error in Script: " & ErrMsg & vbNewLine
  End if
  output = output & "USAGE: " & WScript.ScriptName & " (CheckForUpdate | UpdatePOS)" 
  MsgBox output, vbOKOnly, "Update POS"
end Sub

'***************************************************************************************

Function ArgumentsOK(ByRef Operation)
  Dim argCount
  argcount = WScript.Arguments.Unnamed.Count
  if argCount <> 1 Then
    ArgumentsOK = "No Operation Specified"
  Else
    operation = UCase(WScript.Arguments.Unnamed(0))
    If operation = "CHECKFORUPDATE" Or operation = "UPDATEPOS" Then
      ArgumentsOK = ""
    Else
      ArgumentsOK = "Operation Must be Either CheckForUpdate or UpdatePOS"
    End If
  End If
End Function

'****************************************************************************************

Sub CheckForUpdate()
  Dim Update, i, RemoteLastModifiedDate
  Update = False
  For i = 0 To UBound(SourceFiles)
    If UpdateAvailable(SourceFiles(i), DestDirs(i), RemoteLastModifiedDate) Then
      Update = True
	  Exit For
	End if
  Next 
  If Update then
    MsgBox "An Updated POS is Available!!!" & vbNewLine & _
	       "Be sure to click the UpdatePOS Icon on the Desktop ASAP!", vbOKOnly, "Update POS"
  End if
End Sub

'****************************************************************************************

Sub UpdatePOS()
  Call ForceCScript()
  Dim Update, i, RemoteLastModifiedDate, output
  Update = False
  For i = 0 To UBound(SourceFiles)
    If UpdateAvailable(SourceFiles(i), DestDirs(i), RemoteLastModifiedDate) Then
	  Update = True
      Call DownloadFile(SourceFiles(i), DestDirs(i), RemoteLastModifiedDate)
	End if
  Next 
  If Update then
    output = "POS files have been updated." & vbNewLine & "Enjoy the new updates!"
  Else
    output = "No need for an update." & vbNewLine & "You are using the latest versions!"
  End if
  MsgBox output, vbOKOnly, "Update POS"
End Sub

'****************************************************************************************

Sub DownloadFile(ByVal Filename, ByVal DestDir, ByVal RemoteLastModifiedDate)
  Dim oShell
  Dim Cmd
  Dim SourceFile, DestFile
  WScript.Echo "Downloading " & Filename & " . . ."
  Set oShell = WScript.CreateObject("WScript.Shell")
  SourceFile = SourceDir & Filename
  DestFile = oShell.ExpandEnvironmentStrings(DestDir & "\" & Filename)
  Cmd = oShell.ExpandEnvironmentStrings(ScriptDir & "wget -N " & SourceFile & " -P """ & DestDir & """")
  oShell.Run Cmd, 1, True
  'Cmd = "touch -m -c -t " & Year(RemoteLastModifiedDate) & RightPad(Month(RemoteLastModifiedDate), 2, "0") & _
  '      RightPad(Day(RemoteLastModifiedDate), 2, "0") & RightPad(Hour(RemoteLastModifiedDate), 2, "0") & _
  ' 	RightPad(Minute(RemoteLastModifiedDate), 2, "0") & "." & RightPad(Second(RemoteLastModifiedDate), 2, "0") & " " & _
  '      DestFile
  'oShell.Run Cmd, 1, True
  'MsgBox RemoteLastModifiedDate & vbNewLine & Cmd
  Set oShell = Nothing
End Sub

'****************************************************************************************

'Function RightPad(ByVal strText, ByVal intLen, ByVal chrPad)
'    RightPad = Right( String( intLen, chrPad ) & strText, intLen )
'End Function

'****************************************************************************************

Function UpdateAvailable(ByVal Filename, ByVal DestDir, ByRef RemoteLastModifiedDate)
  dim oShell, filesys, localfile
  Dim LocalLastModifiedDate
  RemoteLastModifiedDate = getRemoteFileDate(SourceDir & Filename)
  If RemoteLastModifiedDate = "" Then
    UpdateAvailable = False
  Else
    Set oShell = WScript.CreateObject("WScript.Shell")
	Filename =  oShell.ExpandEnvironmentStrings(DestDir & "\" & Filename)
	Set oShell = Nothing
    Set filesys = CreateObject("Scripting.FileSystemObject")
    If Not filesys.FileExists(Filename) Then
      UpdateAvailable = True
    Else
      Set localfile = filesys.GetFile(Filename)
	  localLastModifiedDate = localfile.DateLastModified
      UpdateAvailable = DateDiff("s", localLastModifiedDate, RemoteLastModifiedDate) > 2
	  Set localfile = Nothing
    End if
    Set filesys = Nothing
  End if
End Function

'****************************************************************************************

Function getRemoteFileDate(ByVal remoteFilename)
  Const HIDDEN_WINDOW = 0
  Dim strComputer: strComputer = "."
  Dim objWMIService, objStartup, objConfig, objProcess, objShell
  Dim strCommand, intReturn, intProcessID
  Dim colMonitoredProcesses, objLatestProcess
  Dim i
  Dim filesys, readfile
  Dim strOut, Line, Index, LIndex
  
  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
  Set objStartup = objWMIService.Get("Win32_ProcessStartup")
  Set objConfig = objStartup.SpawnInstance_
  objConfig.ShowWindow = HIDDEN_WINDOW
  Set objProcess = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
  strCommand = "cmd /c """ & ScriptDir & "wget --quiet --server-response --spider " & remoteFilename & """ 2> stderr.log"
  Set objShell = WScript.CreateObject("WScript.Shell")
  intReturn = objProcess.Create( strCommand, objShell.CurrentDirectory, objConfig, intProcessID )
  Set objShell = Nothing
  If intReturn <> 0 Then
    Wscript.Echo "Process could not be created." & vbNewLine & "Command line: " & strCommand & vbNewLine & "Return value: " & intReturn
  Else
    'Wscript.Echo "Process created." & vbNewLine & "Command line: " & strCommand & vbNewLine & "Process ID: " & intProcessID
    Set colMonitoredProcesses = objWMIService.ExecNotificationQuery("Select * From __InstanceDeletionEvent Within 1 Where TargetInstance ISA 'Win32_Process'")
    i = 0
    Do Until i = 1
      Set objLatestProcess = colMonitoredProcesses.NextEvent
      If objLatestProcess.TargetInstance.ProcessID = intProcessID Then
        i = 1
      End If
      Set objLatestProcess = Nothing
    Loop
	Set colMonitoredProcesses = Nothing

    set filesys = CreateObject("Scripting.FileSystemObject")
	set readfile = filesys.OpenTextFile("stderr.log", 1, false)
    strOut = ""
    Do While Not readfile.AtEndOfStream
      Line = readfile.ReadLine()
	  Index = InStr(1, Line, "Last-Modified:", 1)
	  If Index <> 0 Then
	    LIndex = InStr(Index + Len("Last-Modified:"), Line, "GMT", 1)
	    Line = Mid(Line, Index + Len("Last-Modified:"), LIndex - Index - Len("Last-Modified:") - 1)
	    strOut = DateFromHTTP(Trim(Line)) & vbNewLine
	    Exit Do
	  End if
    Loop
	readfile.close
	Set readfile = Nothing
	Call filesys.DeleteFile("stderr.log", false)
	Set filesys = Nothing
  End If
  Set objProcess = Nothing
  Set objConfig = Nothing
  Set objStartup = Nothing
  Set objWMIService = Nothing

  getRemoteFileDate = strOut
End Function

'****************************************************************************************

Function DateFromHTTP(ByVal HTTPDate)
'Mon, 24 Sep 2012 14:20:29 GMT
  Dim d, sm, y, h, m, s
  Dim Temp
  d = Mid(HTTPDate, 6, 2)
  sm = Mid(HTTPDate, 9, 3)
  y = Mid(HTTPDate, 13, 4)
  h = Mid(HTTPDate, 18, 2)
  m = Mid(HTTPDate, 21, 2)
  s = Mid(HTTPDate, 24, 2)
  DateFromHTTP = DateAdd("n", -getGMTTimeOffset(), CDate(sm & " " & d & ", " & y & " " & h & ":" & m & ":" & s))
End Function

'****************************************************************************************

Function getGMTTimeOffset()
  Dim oShell
  Dim atb
  Dim offsetMin
  Set oShell = WScript.CreateObject("WScript.Shell") 
  atb = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias" 
  offsetMin = oShell.RegRead(atb)
  Set oShell = Nothing
  getGMTTimeOffset = offsetMin
End Function

'****************************************************************************************

Sub ForceCScript()
  Dim oShell
  Set oShell = WScript.CreateObject("WScript.Shell")
  If Instr(1, WScript.FullName, "CScript", vbTextCompare) = 0 Then
    oShell.Run "cscript """ & WScript.ScriptFullName & """ " & WScript.Arguments.Unnamed(0), 1, False
  	Set oShell = Nothing
    WScript.Quit 0
  End If
  Set oShell = Nothing
End Sub

'****************************************************************************************

Function GetCurrentDir()
  Dim objShell
  Set objShell = WScript.CreateObject("WScript.Shell")
  GetCurrentDir = objShell.CurrentDirectory & "\"
  Set objShell = Nothing
End Function

'****************************************************************************************

Function GetScriptDir()
  Dim objShell, objFSO, objFile
  Set objShell = CreateObject("Wscript.Shell")
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objFile = objFSO.GetFile(Wscript.ScriptFullName)
  GetScriptDir = objFSO.GetParentFolderName(objFile) & "\"
  Set oBjFile = Nothing
  Set objFSO = Nothing
  Set objShell = Nothing
End Function
