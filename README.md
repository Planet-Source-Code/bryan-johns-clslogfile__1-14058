<div align="center">

## clsLogFile


</div>

### Description

This class module is very useful for keeping a standardized, formatted, event/error log for any application that might need one.
 
### More Info
 
See the comments in the code. Most should be self explanatory. The DaysToKeep property might be a little obscure. It's an integer value that tells the object how many days of log entries to keep when purging old entries during the Class Terminate process.

Upon instantiating the object, set the DaysToKeep property to some integer value. The higher the value the further back in time the error.log entries will go. There are two basic usages of this class module. One is a generic error handler using the SimpleError method. This method should be called from any error handling routines you may have in your code. It takes as optional parameters the sub name and the form name of the code where it's being called. This is useful in tracking down where an error happened if a user calls for tech support. The other usage can be used to log events in code that you for whatever reason wish to log. For example, if you want to keep track of when a database file is opened and closed. In this usage you'd call the WriteLog method. This method takes several parameters. A string to hold the desired message, an optional parameter for the sub and for the form name, as well as an optional parameter to tell the object if this is a new entry or part of an existing one. For that last parameter, if true the message is handled like a new entry, enclosed in a block of *'s to make it easy to pick out from other messages. If false then it's treated as a continuation of the previous message. This allows multi-line messages.

Log entries are writte to a file called error.log which is stored in the application path. It is formatted with each individual log entry seperated by *'s and a header line showing which file or form and procedure generated the error.

If the DaysToKeep property is set too high it could result in a very large log file.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bryan Johns](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bryan-johns.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bryan-johns-clslogfile__1-14058/archive/master.zip)





### Source Code

```
Option Explicit
'local variable(s) to hold property value(s)
Private mvarDaysToKeep As Integer 'local copy
Private Const File As String = "classLogFile"
Public Property Let DaysToKeep(ByVal vData As Integer)
  mvarDaysToKeep = vData
End Property
Public Property Get DaysToKeep() As Integer
  DaysToKeep = mvarDaysToKeep
End Property
Public Sub WriteLog(lstrMessage As String, Optional lstrProc As String, Optional lstrFile As String, Optional lboolNewEntry As Boolean)
'**************************************************************
'* procedure to write out log entries
'* it accepts the following parameters:
'*   lstrMessage (String containing the message to be logged)
'*   lstrProc (optional string containing the procedure that
'*     generated the log entry)
'*   lstrFile (optional string containing the file that
'*     contains the procedure that generated the log entry)
'*   lboolNewEntry (optional boolean to force the procedure
'*     to treat this entry as a new entry thereby adding
'*     the entry separation formatting)
'***************************************************************
  Dim lstrMyDate As String
  Dim lstrMyTime As String
  Dim lstrFileName As String
  Dim lintFileNum As Integer
  Dim lstrLogMessage As String
  Dim msg As String
  Const SubName = "Public Sub oError.WriteLog(lstrMessage As String, Optional lstrProc As String, Optional lstrFile As String, Optional lboolNewEntry As Boolean)"
  On Error GoTo Error
  ' get a free file number for the error.log file
  lintFileNum = FreeFile
  ' assign the file name
  lstrFileName = App.Path & "\error.log"
  ' open the log file
  Open lstrFileName For Append As lintFileNum
  ' format and initialize the date and time variables
  lstrMyDate = Format(Date, "mmm dd yyyy")
  lstrMyTime = Format(Time, "hh:mm:ss AMPM")
  If lboolNewEntry = True Then
    ' write the top boundary of the log entry.
    lstrLogMessage = lstrMyDate & " " & lstrMyTime & " ********************************************************************************** "
    Print #lintFileNum, lstrLogMessage
    If Len(lstrFile) > 0 Then ' write the file
      lstrLogMessage = lstrMyDate & " " & lstrMyTime & " *** File: " & lstrFile
    Else
      lstrLogMessage = lstrMyDate & " " & lstrMyTime & " *** File: Not Supplied"
    End If
    If Len(lstrProc) > 0 Then ' write the procedure
      lstrLogMessage = lstrLogMessage & " ***** " & " Procedure: " & lstrProc
    Else
      lstrLogMessage = lstrLogMessage & " ***** " & " Procedure: Not Supplied"
    End If
    Print #lintFileNum, lstrLogMessage
  End If
  ' write the log entry
  lstrLogMessage = lstrMyDate & " " & lstrMyTime & " *** " & lstrMessage
  Print #lintFileNum, lstrLogMessage
  If lstrMessage = "Normal Exit" Then
    ' write the bottom boundary of the log entry.
    lstrLogMessage = lstrMyDate & " " & lstrMyTime & " ********************************************************************************** "
    Print #lintFileNum, lstrLogMessage
  End If
  'close the log file
  Close lintFileNum
  Exit Sub
Error:
  msg = "Error in creating or editing the error.log file." & vbCrLf
  msg = msg & "Error: " & Err.Number & " - " & Err.Description & vbCrLf
  msg = msg & "Program File: " & File & "Procedure: " & SubName
  MsgBox msg, vbCritical
End Sub
Private Sub RemoveOldLogEntries(Days As Integer)
'*************************************************************
'* RemoveOldLogEntries is a procedure that, as it's name
'* implies parses thru the lines in the error log file created
'* in the above oError.WriteLog procedure and removes entries
'* past an number of days specified at the time this procedure
'* is called
'* It accepts the following parameters:
'*   Days (an integer that specifies the number of days
'*     beyond which to delete the log entries)
'*************************************************************
  Dim lstrInFileName, lstrOutFileName As String
  Dim lstrLogEntry, lstrEntryDate As String
  Dim lintInFileNum, lintOutFileNum As Integer
  Const SubName = "Private Sub RemoveOldLogEntries(Days As Integer)"
  On Error GoTo Error
  WriteLog "Removing log entries greater than " & Str(Days) & " days old.", SubName, File, False
  ' assign the file name
  lstrInFileName = App.Path & "\error.log"
  lstrOutFileName = App.Path & "\error.tmp"
  If Dir(lstrInFileName) = "error.log" Then
    ' get a free file number for the error.log file
    lintInFileNum = FreeFile
    ' open the error.log file for reading and the error.tmp file for writing
    Open lstrInFileName For Input As lintInFileNum
    lintOutFileNum = FreeFile
    Open lstrOutFileName For Append As lintOutFileNum
    Do While Not EOF(lintInFileNum)
      Line Input #lintInFileNum, lstrLogEntry  ' Read line into variable.
      lstrEntryDate = Left(lstrLogEntry, 11)
      If DateDiff("d", lstrEntryDate, Now) <= Days Then
        Print #lintOutFileNum, lstrLogEntry
        Exit Do
      End If
RecoverFromError:
    On Error GoTo Error:
    Loop
    Do While Not EOF(1)
      Line Input #lintInFileNum, lstrLogEntry
      Print #lintOutFileNum, lstrLogEntry
    Loop
    Close #lintInFileNum  ' Close file.
    Close #lintOutFileNum
    Kill lstrInFileName
    Name lstrOutFileName As lstrInFileName
  End If
  Exit Sub
Error:
  If Err.Number = "13" Then
    GoTo RecoverFromError
  End If
  MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub
Public Sub SimpleError(Optional SubName As String, Optional FormName As String)
  Dim msg As String
  If Len(SubName) = 0 Then SubName = "Unspecified"
  If Len(FormName) = 0 Then SubName = "Unspecified"
  msg = "Error: " & Err.Number & " - " & Err.Description
  MsgBox msg, vbCritical
  WriteLog msg, SubName, FormName, True
End Sub
Private Sub Class_Initialize()
  WriteLog App.EXEName & " Started", "Private Sub Class_Initialize()", File, True
  DaysToKeep = 1
End Sub
Private Sub Class_Terminate()
  WriteLog "Terminating LogFile Object", "Private Sub Class_Terminate()", File, True
  RemoveOldLogEntries DaysToKeep
  WriteLog "Normal Exit", "Private Sub Class_Terminate()", File, True
End Sub
```

