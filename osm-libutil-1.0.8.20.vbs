'********************************************************************
'* Filename:  libutil-1.0.8.vbs
'* Author:    Randoll REVERS
'* Version:   1.0.8 (build 20)
'* Date:      2006.07.11
'* Purpose:   Reusable methods and properties 
'*
'* RevInfo:   Updated last on 2011.12.17/RRR
'********************************************************************
OPTION EXPLICIT

   Const INVALID = -1
   Const ERRO = 0
   
   Const FSO_READ = 1
   Const FSO_WRITE = 2
   Const FSO_APPENDING = 8
   
   Const P_DIR_OK_FILE_OK = 0
   Const P_FILE_OK = 1
   Const P_DIR_OK = 2
   Const P_DIR_OK_FILE_NOT = 3
   Const P_DIR_NOT_FILE_NOT = 4   
   Const P_DIR_NOT = 4   
   Const P_ADD = 10
   Const P_REMOVE = 11
   
   Const LOG_NONE = 20
   Const LOG_ON = 21
   Const LOG_OFF = 22
   
   Const S_NOEXT = 0
   Const S_EXT = 1
     
   
   dim FSO : Set FSO = CreateObject( "Scripting.FileSystemObject")
   dim WSO : Set WSO = CreateObject( "WScript.Shell")
   dim SO  : Set SO  = GetObject( "winmgmts://./")
   
   dim mLogFile_      : mLogFile_  = null
   dim mLogState_     : mLogState_ = LOG_NONE
   dim mLogFmt_       : mLogFmt_ = ""
   dim mLogTimestamp_ : mLogTimestamp_ = true
   dim mLogDTFmt_     : mLogDTFmt_ = "YYYY-mm-dd HH:MM:ss "
   
   dim mCmdl_     : mCmdl_ = null
   dim mCmdlErr_  : mCmdlErr_ = false
   
   
   '**
   '* Method:       ExitProcess
   '* Ver/Author:   1.0.1/RRR
   '* Purpose:      Exit process with a message and return code.
   '* Date:         2007.06.18
   '* Rev-Info:     (see change history below)
   '*
   Sub ExitProcess( ByVal doneMsg, ByVal rCode)
      If Not isNull( doneMsg) Then
         doneMsg = Right( doneMsg, Len(doneMsg)-1)
         If Not Left( doneMsg, 1)= "-" Then doneMsg = "Done. ("&getScriptName( S_EXT)&")"
      Else
         doneMsg = "Done. ("&getScriptName( S_EXT)&")"
      End If
      msgLog doneMsg
      WScript.Quit( rCode)
   End Sub
   
   
   '**
   '* Method:  failedCmdl
   '* Author:  RRR
   '* Date:    2007.06.27
   '* Purpose: Display command-line usage as exceptions occur
   '*
   Sub failedCmdl()
      mCmdlErr_ = true
      verifyCmdl getArgCount()+1, iif( isNull( mCmdl_), "<verify script documentation>", mCmdl_)
   End Sub


   '**
   '* Method:  fileVer
   '* Author:  RRR
   '* Date:    2006.04.13
   '* Purpose: Retrieve file version
   '*
   '* RevInfo: (see change history below)
   '*
   Function fileVer( ByVal file)
      On Error Resume Next 
      dim fst, tls_log_, ver : ver = -1

      msgLog "FileVer(): "&file
      
      '-- turn off logging
      tls_log_ = mLogState_
      mLogState_ = LOG_OFF 
      fst = Path( file, false)
      mLogState_ = tls_log_

      '-- get file version
      If fst = P_DIR_OK_FILE_OK Or fst = P_FILE_OK Then
         ver = FSO.GetFileVersion( file)
      End If     

      '-- validation 
      If ver > -1 Then
         msgLog "!version is "&ver&"."
      Else
         msgLog "!Could not retrieved version."
      End If
      fileVer = ver
   End Function 'fileVer()


   '**
   '* Method:       formatBytes
   '* Ver/Author:   1.0.0/RRR
   '* Purpose:      Returns compact formatted string of byte size
   '* Date:         2011.05.13
   '*
   Function formatBytes( byVal v)
      dim s,b : s = len(v)

      If s>12 then
         b = Round(v/(1024*1024*1024*1024),2)&" TB"
      Else 
         If s>9 then
            b = Round(v/(1024*1024*1024),2)&" GB"
         Else
            If s>6 then
               b = Round(v/(1024*1024),2)&" MB"
            Else
               If s>3 then
                  b = Round(v/1024,2)&" KB"
               Else
                  b = v&" bytes"
               End If
            End If
         End If
      End If

      formatBytes = b 
   End Function 'formatBytes()


   '**
   '* Method:  iif
   '* Author:  RRR
   '* Date:    2006.07.14
   '* Purpose: Simulates conditional statements, 
   '*          (expression) ? true : false
   '*
   '* example: result = iff( x > 5, "Value exceeded!", "Value validated!")
   '*          result = iff( x = true, do_post, do_quit)
   '*
   Function iif( ByVal exp, ByVal t, ByVal f)
      If exp = true Then 
         iif = t
         Exit Function
      End If
      iif = f
   End Function 'iif()
      
   
   '**
   '* Method:       getArgCount
   '* Ver/Author:   1.0.0/RRR
   '* Purpose:      Returns command line arguments count
   '* Date:         2007.06.20
   '*
   Function getArgCount()
      getArgCount = WScript.Arguments.Count
   End Function 'getArguments
   
   '**
   '* Method:       getArgNamed
   '* Ver/Author:   1.0.0/RRR
   '* Purpose:      Returns commandline named arguments
   '* Date:         2007.06.20
   '*
   Function getArgNamed( ByVal argName_)
      getArgNamed = iif( isEmpty(WScript.Arguments.Named( argName_)), _
         null, WScript.Arguments.Named( argName_))
   End Function 'getArgument()


   '**
   '* Method:  getCurrentPath()
   '* Author:  RRR
   '* Date:    2006.02.21
   '* Purpose: Returns the current path where the script is running
   '*
   Function getCurrentPath()
      dim pathstr:pathstr=WScript.scriptfullname
      getCurrentPath=left(pathstr,len(pathstr)-len(WScript.scriptname)-1)
   End Function 'getCurrentPath()
   
   
   '**
   '* Method:  getDate
   '* Author:  RRR
   '* Date:    2006.03.06
   '* Purpose: Returns current Date (i.e. 20060306)
   '* Rev:     2010.12.22/RRR, use getFDate()
   Function getDate()
      getDate = getFDate( "YYYYmmdd")
   End Function 'getDate()
   
   
   '**
   '* Method:  getFDate
   '* Author:  RRR
   '* Date:    2010.12.22
   '* Purpose: Returns current Date formatted with pattern
   Function getFDate( byVal s)
      s = replace( s, "YYYY", cstr(year(now())))
      s = replace( s, "mm", right("00" & cstr(month(now())),2))
      s = replace( s, "dd", right("00" & cstr(day(now())),2)) 
      getFDate = s 
   End Function 'getFDate()


   '**
   '* Method:  getEnvar
   '* Author:  RRR
   '* Date:    2006.03.02
   '* Purpose: Retrieve Process/System level environment variable
   '*
   Function getEnvar( ByVal var)
      On Error Resume Next
      dim val : val = WSO.ExpandEnvironmentStrings("%"&var&"%")
      If Err.Number <> 0 Then
         msgLog "GetEnvar(): %"&var&"% could not be found!"
         val = null
      Else
         msgLog "GetEnvar(): %"&var&"% = "&val
      End If

      getEnvar = val
   End Function 'getEnvar()
   
   
   '**
   '* Method:  getScriptName
   '* Author:  RRR
   '* Date:    2007.06.27
   '* Purpose: Returns the current script filename 
   '*
   Function getScriptName( byVal ext)
      dim name : name = WScript.ScriptName
      If ext = S_NOEXT Then
         name = left(name,len(name)-4)
      End If
      getScriptName = name
   End Function

   
   '**
   '* Method:  getTime
   '* Author:  RRR
   '* Date:    2006.03.06
   '* Purpose: Returns current Time (i.e. 114223)
   '* Rev:     2010.12.22/RRR, use getFTime()
   Function getTime()
      getTime = getFTime( "HHMMss")
   End Function 'getTime()
   
   
   '**
   '* Method:  getFTime
   '* Author:  RRR
   '* Date:    2010.12.22
   '* Purpose: Returns current Time formatted with pattern
   Function getFTime( byVal s)
      s = replace( s, "HH", right("00"&cstr(hour(now())),2))
      s = replace( s, "MM", right("00"&cstr(minute(now())),2))
      s = replace( s, "ss", right("00"&cstr(second(now())),2))
      getFTime = s
   End Function 'getFTime()


   '**
   '* Method:  logger
   '* Author:  RRR
   '* Date:    2006.04.13
   '* Purpose: Create a log file. 
   '*
   '* example: logger "<path+filename>", file and path to attribute to the log
   '*          logger "#on", write to log file
   '*          logger "#off", don't write anything to the log anymore
   '*          logger "#reset", override the log everytime
   '*          logger "#>>", indentation forward 
   '*          logger "#<<", indentation backward
   '*
   '* RevInfo: (see change history below)
   '*
   Sub logger( ByVal log)
      On Error Resume Next
      dim fst, feature 
      log = LCase( log)
      feature = log
      
      
      '-- validate log file refence
      If mLogFile_ = "" Or isEmpty( mLogFile_) Then mLogFile_ = null
      
      
      '-- separate feature from filename
      If Left( log, 1) = "#" Then 
      
         '-- feature manipulation
         Select Case feature
            case "#>>"
               'toggle indentation on (support multi-level)
               mLogFmt_ = mLogFmt_ + "~   "
            case "#<<"
               'toggle indentation off (support multi-level)
               If len( mLogFmt_) > 0 Then mLogFmt_ = left( mLogFmt_, len( mLogFmt_)-4)
            case "#on"
               mLogState_ = LOG_OFF
               fst = Path( mLogFile_, false)
               If Not isNull( mLogFile_) And _ 
                  fst = P_DIR_OK_FILE_OK Or _
                  fst = P_FILE_OK Then 
                  mLogState_ = LOG_ON 'turn on logging
               Else
                  mLogState_ = LOG_NONE 'turn on logging, but display on screen
               End If
            case "#off"
               mLogState_ = LOG_OFF 'turn off logging
            case "#reset"
               If FSO.FileExists( mLogFile_) And Not isNull( mLogFile_) Then 
                  dim gf : set gf = FSO.GetFile( mLogFile_)
                  gf.delete
                  'If Err.Number = 0 Then
                  '   msgLog "Log file """&UCase( mLogFile_)&""" has been reset."
                  'Else
                  '   msgLog "Unable to reset log file: "&iif( isNull( mLogFile_), "NONE", mLogFile_)
                  'End If
                  If Err.Number <> 0 Then msgLog "Unable to reset log file: "&iif( isNull( mLogFile_), "NONE", mLogFile_)
               End If
            case else
               msgLog "Logger(): Setting logger feature not failed!"
         End Select
      Else
      
         '-- log file manipulation
         dim tls_log_ : tls_log_ = mLogState_
         If len( log) > 1 Then
            mLogState_ = LOG_OFF
            fst = Path( log, true)
            If fst = P_DIR_OK_FILE_OK Or _
               fst = P_DIR_OK_FILE_NOT Or _
               fst = P_FILE_OK Or _
               fst = P_FILE_NOT Then 
               mLogFile_ = log
               mLogState_ = LOG_ON
            Else
               mLogFile_ = null
            End If 
            mLogState_ = tls_log_
         End If
      End If
   End Sub 'logger()
   

   '**
   '* Method:  Lpad
   '* Author:  RRR
   '* Date:    2011.09.29
   '* Purpose: Returns string with provided padding
   Function Lpad( byVal s, byVal p, byVal l)
      Lpad = Rstr(l-len(s),p)&s
   End Function


   '**
   '* Method:  msgLog
   '* Author:  RRR
   '* Date:    2006.03.27
   '* Purpose: Applend to log routine 
   '*          - Close log file after usage to avoid access conflict
   '* Features:
   '*          - '""' or ':' -- apply no timestamp if string starts with special char
   '*          - '!' -- apply special formatting for sub-routine messages
   '*          - '@' -- apply simultaneous logging to string, screen+log
   '*          - (note order: ":!@<text>")
   '* RevInfo: (see change history below)
   '*
   Sub msgLog( ByVal strText) 
      On Error Resume Next 
      Dim wf, fmtTimestamp, fmtIndentSub, fmtScrEna
      
      '-- validate log file status 
      If isNull( strText) Or mLogState_ = LOG_OFF Then Exit Sub 
      
      '-- get output stream 
      If isNull( mLogFile_) Then 
         Set wf = WScript.StdOut 
      Else 
         mLogState_ = LOG_ON
         If Not FSO.FileExists( mLogFile_) Then FSO.CreateTextFile( mLogFile_) 
         Set wf = FSO.OpenTextFile( mLogFile_, FSO_APPENDING) 
         If Err.Number <> 0 Then 
            mLogState_ = LOG_NONE
            mLogFile_ = null 
         End If
      End If
                 
      '-- apply feature: "" or ":", do or don't apply timestamp
      If strText = "" Or left( strText, 1) = ":" Then 
         If Left( strText, 1) = ":" Then strText = Right( strText, Len( strText)-1) 'remove special char
         fmtTimestamp = ""
      Else
         fmtTimestamp = getFDate(mLogDTFmt_)
         fmtTimestamp = getFTime(fmtTimestamp)
      End If
      '-- apply feature: "!", indent sub-message
      If Left( strText, 1) = "!" Then
         strText = Right( strText, Len( strText)-1) 'remove special char
         fmtIndentSub = "   ["
      Else
         fmtIndentSub = ""
      End If 
      '-- apply feature: "@", simultaneously output to screen and log file
      If left( strText, 1) = "@" Then
         fmtScrEna = "" 
         strText = Right( strText, Len( strText)-1) 'remove special char
         If mLogState_ = LOG_ON Then 
            fmtScrEna = iif( fmtTimestamp = "", "", "((( ")
            WScript.StdOut.WriteLine( " "&mLogFmt_&fmtIndentSub&strText& _
            iif(fmtIndentSub="","","]"))
         End If
      End If
      '-- apply feature: "#timestamp", print timestamp only
      If strText = "#timestamp" And mLogTimestamp_ = false Then 
         strText = "- -------------------"&vbCRLF& _
                   "- "&fmtTimestamp&vbCRLF& _
                   "- -------------------"
         fmtTimestamp = ""
      Else
         fmtTimestamp = iif(fmtTimestamp="","",iif(mLogTimestamp_=false,"- ",fmtTimestamp&"- "))
      End If
      
      '-- logging message
      wf.WriteLine( _
         fmtTimestamp& _ 
         fmtScrEna&mLogFmt_& _
         fmtIndentSub&strText& _
         iif(fmtIndentSub="","","]")& _
         iif(fmtScrEna="",""," )))"))
      If mLogState_ = LOG_ON Then wf.Close
   End Sub 'msgLog()
   
   
   '**
   '* Method:  Path
   '* Author:  RRR
   '* Date:    2006.03.10
   '* Purpose: Verify Filesystem path. 
   '*          Limitation: Make sure the last folder has a backslash '\'
   '*          even if there is no file specified after the backslash.
   '*          * If a file is specified and does not exist, and bCreate is true
   '*          the folder path will be create, but the return will be false
   '*
   '* RevInfo: (see change history below)
   '*
   Function Path( ByVal url, ByVal bCreate)
      On Error Resume Next

      dim statusFolder : statusFolder = false
      dim statusFile : statusFile = false
      dim status : status = false
      dim splSize : splSize = 0
      dim splPath : splPath = Split( url, "\", -1, vbTextCompare)
      dim currentPath : currentPath = ""

      msgLog "Path(): validation..."""&url&""""

      '-- check Folder(s) Existence 
      dim iter : For iter = LBound( splPath) To UBound( splPath)-1
         splSize = splSize + 1
         If iter = 0 Then 
            currentPath = splPath( iter)
         Else 
            currentPath = currentPath&"\"&splPath( iter)
         End If 
         If FSO.FolderExists( currentPath) Then
            statusFolder = true 
         Else 
            If bCreate = true Then
               FSO.CreateFolder( currentPath)
               If Err.Number <> 0 Then
                  statusFolder = false
                  Exit For
               End If
               msgLog "!Created folder: "&currentPath&"\"
               statusFolder = true 
            Else
               statusFolder = false 
            End If
         End If         
      Next
            
      '-- check Specified File Existence 
      If Len( url) <> Len( currentPath&"\") Then 
         If FSO.FileExists( url) Then
            statusFile = true
         Else
            statusFile = false
         End If
      Else
         statusFile = null 'there is no file in path
      End If

      '-- validation 
      If statusFolder = false Then
         If splSize < 1 And statusFile = true Then
            msgLog "!File is present."
            status = P_FILE_OK
         Else
            If bCreate Then 
               msgLog "!Unable to create path."
            Else
               msgLog "!Unable to locate path."
            End If
            status = P_DIR_NOT_FILE_NOT
         End If
      Else
         If not isNull( statusFile) Then
            If statusFile = false Then
               msgLog "!Path is present, but file could not be located."
               status = P_DIR_OK_FILE_NOT
            Else
               msgLog "!Path and file are present."
               status = P_DIR_OK_FILE_OK
            End If
         Else
            status = P_DIR_OK
            msgLog "!Path is present."
         End If
      End If
      Path = status
   End Function 'Path()
   
   
   '**
   '* Method:  Path_
   '* Author:  RRR
   '* Date:    2011.09.22
   '* Purpose: (same Path() without logging).
   Function Path_( ByVal url, ByVal bCreate)
      logger "#off"
      Path_ = Path( url, bCreate)
      logger "#on"
   End Function
   
   
   '**
   '* Method:  Rpad
   '* Author:  RRR
   '* Date:    2011.09.29
   '* Purpose: Returns string with provided padding
   Function Rpad( byVal s, byVal p, byVal l)
     Rpad = s&Rstr(l-len(s),p)
   End Function


   '**
   '* Method:  Rstr
   '* Author:  RRR
   '* Date:    2011.09.29
   '* Purpose: Returns string after repeating and parsing it "n" times
   Function Rstr( byVal n, byVal s)
      dim i, t
      For i = 0 to n
         t = t & s
      Next
      Rstr = t
   End Function
  
   
   '**
   '* Method:  removeEnVar
   '* Author:  RRR
   '* Date:    2006.03.02
   '* Purpose: Remove Process level environment variable
   '*
   Sub removeEnvar( ByVal var)
      dim objEnv : set objEnv = WSO.Environment("Process") 
      msgLog "RemoveEnvar(): Clear environment variable ("&var&")."
      objEnv.Remove( var)
   End Sub 'removeEnvar()

   
   '**
   '* Method:  setEnvar
   '* Author:  RRR
   '* Date:    2006.03.02
   '* Purpose: Set Process level environment variable
   '*          Alternative: create/append a file that can be run to set local var
   '*
   Sub setEnvar( ByVal var, ByVal val)
      dim objEnv : set objEnv = WSO.Environment("Process") 
      objEnv( var) = val 
      msgLog "SetEnvar(): "&var&"="&val
   End Sub 'setEnvar()
   
   
   '**
   '* Method:  setPathEnvar
   '* Author:  RRR
   '* Date:    2007.04.09
   '* Purpose: Append/Remove "string" in %PATH%
   '*
   Sub setPathEnvar( ByVal pth, ByVal act)
      On Error Resume next
      dim newPath, curPath, action, msg, pos, sep, key
      
      key = "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\Path"
      curPath = WSO.RegRead( key)
      If Err.Number <> 0 Then
         msg = "Couldn't acces Registry key with %PATH% value."
         Exit Sub
      Else      
         If InStr( UCase(curPath), UCase(pth)) > 0 Then
         Else
         End If
         Select Case (act)
            Case P_ADD
               action = "ADD"
               If InStr( UCase(curPath), UCase(pth)) > 0 Then
                  msg = """String"" is already in %PATH%"
               Else 'Add "string" in %PATH%
                  newPath = curPath&";"&pth
                  msg = newPath
                  WSO.RegWrite key, newPath, "REG_EXPAND_SZ"
                  setEnvar "PATH", newPath
               End If
            Case P_REMOVE 'Remove "string" in %PATH%
               action = "REMOVE"
               pos = InStr( UCase(curPath), UCase(pth))
               If pos > 0 Then
                  sep = Mid( curPath, pos-1, 1) 'Check seperator before
                  If IsEmpty( sep) or sep <> ";" Then 
                     sep = Mid( curPath, pos+len(pth)+1, 1) 'Check seperator after
                     If IsEmpty( sep) or sep <> ";" Then 
                        newPath = Replace( UCase(curPath), UCase(pth)&";", "")
                     Else
                        newPath = Replace( UCase(curPath), UCase(pth), "")
                     End If
                  Else
                     newPath = Replace( UCase(curPath), ";"&UCase(pth), "")
                  End If
                  msg = newPath
                  WSO.RegWrite key, newPath, "REG_EXPAND_SZ"
                  setEnvar "PATH", newPath
               Else
                  msg = """String"" is not in %PATH%"
               End If
            Case Else
               msg = "Invalid usage! Action requsted unknown."
         End Select
         
         msgLog "SetPathEnvar(): "&action&" """&pth&""""
         msgLog "!Before: "&curPath
         msgLog "!After:  "&msg
      End If
   End Sub 'SetPathEnvar() 
   
   
   '**
   '* Method:  setRunOnce
   '* Author:  RRR
   '* Date:    2006.03.29
   '* Purpose: Add parameters to RunOnce in the regsitry
   '*
   Sub setRunOnce( ByVal key, ByVal val)
      WSO.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce\"&key, val, "REG_SZ"
      msgLog "SetRunOnce(): Key["&key&"], Value["&val&"]"
   End Sub 'setRunOnce()


   '**
   '* Method:  setAutoTimestamp
   '* Author:  RRR
   '* Date:    2011.12.17
   '* Purpose: Turn on and automatic time stamping in log
   '*
   Sub setAutoTimestamp( ByVal state)
      mLogTimestamp_ = state
   End Sub 'setAutoTimestamp()
   
   
   '**
   '* Method:  verifyCmdl
   '* Author:  RRR
   '* Date:    2007.06.27
   '* Purpose: Simplifed way to check required number of switches were provided
   '*
   Sub verifyCmdl( ByVal reqN_, ByVal useMsg_)
      mCmdl_ = useMsg_
      If getArgCount() < reqN_ Then
         WScript.StdOut.writeLine vbCrlf&" Problem occured with command-line switches."
         WScript.StdOut.writeLine vbCrlf&"   usage: "&getScriptName( S_EXT)&" "&useMsg_
         logger "#off"
         ExitProcess Null, 1
      End If
   End Sub

   
''' HISTORY '''

' 20060607 0.0.0.0  Added support for "" to msgLog().
' 20060707 0.0.0.0  Bugfix, Prevent logging in msgLog() while setting mLogFile_ because
'                   of the use of getEnvar() causing StackOverFlow.
' 20060724 1.0.0.0  Ported methods and properties from common-loe-1.0.2.wsc:
'                      FSO (for Scripting.FileSystemObject object)
'                      WSO (for WScript.Shell object)
'                      SO (for current system's WMI container)
'                      Path()
'                      fileVer()
'                      iif() - from Class BkJob, (bkjobs.vbs draft)
'                      getDate()
'                      getEnvar(), returns environment variables
'                      getTime()
'                      logger()
'                      msgLog()
'                      removeEnvar(), remove environment variables
'                      setEnvar(), set environment variables
'                      setRunOnce(), set RunOnce in registry
' 20060725 1.0.0.1  Ported getCurrentPath() from WSC framework usage.
'                   Optimized Path() for better returned status.
'                   Added support for optimized Path() status into filever() and logger().
' 20060725 0.0.0.0  logger(): Simplified logging mechanism to work even with service accounts.
'                             Added support for optimized Path() status.
' 20060725 0.0.0.0  Path(): Optimized for better returned status.
' 20060725 0.0.0.0  filever(): Added support for optimized Path() status.
' 20060801 0.0.0.0  msgLog(): Added indentation for sub/related-infos.
' 20060802 0.0.0.0  logger(): Added better support for features.
'                             Added indentation features.
' 20061012 1.0.0.2  Fix bug in Path, null value handler for file status check.
' 20061129 1.0.0.3  Optimized logger to be able to reset current log file.
' 20061207 2.0.0.0  Merge libutil-1.0.1 with libloe-1.0.1.
' 20070404 1.0.2.4  Unmerge from libutil-2.0.0-vbs.
' 20070409 1.0.3.5  Implement setPathEnvar(), to add/remove path strings in the %PATH%.
' 20070608 1.0.3.6  Optimized msgLog() to print logging on STDOUT when no logfile is specified.
' 20070615 1.0.4.7  Optimized msgLog()'s algorithm and add feature "@<txt>" to simultaneous print
'                   onto STDOUT and into the log file.
' 20070618 1.0.4.8  Added ExitProcess(), to exit the current process with a message and return code.
' 20070620 1.0.4.9  Optimized logger() to better differentiate features from log filenames.
' 20070620 1.0.5.10 Added getArgCount(), returns command line arguments count.
' 20070620 1.0.5.11 Added getArgNamed(), returns commandline named argument (passed) or null.
' 20070627 1.0.6.12 Added verifyCmdl() to check required number of switches were provided.
' 20070627 1.0.6.13 Added failedCmdl() to handle exception with the cmdline switches.
' 20090603 1.0.6.14 Optimized ExitProcess(), added "-" feature to supress done message.
' 20101222 1.0.7.15 Added getFDate( "YYYY.mm.dd") add support for partern-based date formatting.
' 20101222 1.0.7.16 Added getFTime( "HHMMss") add support for partern-based time formatting.
' 20110513 1.0.8.17 Added formatBytes(), returns compact formatted string of byte size.
' 20110922 1.0.8.18 Added Path_() without logging.
' 20110929 1.0.8.19 Added Lpad(), Rpad() and Rstr().
' 20111217 1.0.8.20 Added mLogTimestamp_ and setAutoTimestamp() to remove automatic time stamping.
'                   Optimized timestamp formatting in msgLog with mLogDTFmt_.
'                   Added new feature to msgLog ("#timestamp"), for time stamping logs.
'                   Removed log reset message.
