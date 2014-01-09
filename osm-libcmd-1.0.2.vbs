'**********************************************************
'* Filename:  libcmd-1.0.2.vbs
'* Author:    Randoll REVERS
'* Version:   1.0.2 (build 10) | requires libutil-1.0.4.14 or later
'* Date:      2006.07.11
'* Purpose:   Systems Common Commands
'*
'* RevInfo:   
'**********************************************************

   Const L_SRV_RUNNING = 1
   Const L_SRV_STOPPED = 2
   Const L_SRV_DISABLED = 3
   Const L_SRV_AUTO = 4
   Const L_SRV_MANUAL = 5
   
   Const L_STD_WAIT = 10000


   Class SYSCmd
               
      '**
      '* Set default values 
      '*
      Public default Sub Class_Initialize()
         
      End Sub


      '**
      '* Cleanup routine
      '*
      Private Sub Class_Terminate()
         
      End Sub


      '**
      '* Method:       copyFile
      '* Ver/Author:   1.0.1/RRR
      '* Date:         2006.04.14
      '* Purpose:      Copy a file
      '*               Improved readability. Exception Handler.
      '*
      Public Function copyFile( src, trg)
         On Error Resume Next
   
         dim status : status = false
   
         If src = trg Then 
            msgLog "CopyFile(): File was not copied. <source> = <target> path."
            copyFile = true
            Exit Function
         End If 
         
         If FSO.FileExists( src) Then
            FSO.CopyFile src, trg, true
            If Err.Number <> 0 Then
               status = false
            Else
               status = true
            End If
         End If
   
         ''' Validate '''
         If status = false Then 
            msgLog "CopyFile(): File was not copied."
         Else
            msgLog "CopyFile(): File copied."
         End If
            msgLog "   [source] "&src
            msgLog "   [target] "&trg
         copyFile = status
      End Function
   
   
      '**
      '* Method:       copyFolder
      '* Ver/Author:   1.0.1/RRR
      '* Date:         2006.04.14
      '* Purpose:      Copy a folder
      '*               Improved readability. Exception Handler.
      '*
      Public Function copyFolder( src, trg)
         On Error Resume Next
   
         dim status : status = false
   
         If src = trg Then 
            msgLog "CopyFolder(): Folder was not copied. <source> = <target> path."
            copyFolder = true
            Exit Function
         End If 
         
         If FSO.FolderExists( src) Then
            FSO.CopyFolder src, trg, true
            If Err.Number <> 0 Then
               status = false
            Else
               status = true
            End If
         End If
   
         ''' Validate '''
         If status = false Then 
            msgLog "CopyFolder(): Folder was not copied."
         Else
            msgLog "CopyFolder(): Folder copied."
         End If
         msgLog "   [source] "&src
         msgLog "   [target] "&trg
         copyFolder = status
      End Function 
   
   
      '**
      '* Method:       deleteFile
      '* Ver/Author:   1.0.0/RRR
      '* Date:         2006.05.03
      '* Purpose:      Delete File
      '*
      Public Function deleteFile( file)
         On Error Resume Next
         
         dim status : status = false
         dim msg, gf 
         
         If FSO.FileExists( file) Then
            set gf = FSO.GetFile( file)
            gf.delete
            If Err.Number <> 0 Then
               msg = "(Error: "&iif( Err.Description = "", "?", Err.Description)&")"
               status = false
            Else
               status = true
            End If
         Else
            msg = "(File not found!)"
         End If
         
         ''' Validate '''
         If status = false Then 
            msgLog "DeleteFile(): File was not deleted. "&msg
         Else
            msgLog "DeleteFile(): File deleted."
         End If
         msgLog "   [file] "&file
         
         gf.Close()
         set gf = Nothing
         deleteFile = status
      End Function 
      
      
      '**
      '* Method:       deleteFolder
      '* Ver/Author:   1.0.0/RRR
      '* Date:         2006.05.03
      '* Purpose:      Delete Folder
      '*
      Public Function deleteFolder( fld)
         On Error Resume Next
         
         dim status : status = false
         dim msg 
         
         If FSO.FolderExists( fld) Then
            FSO.DeleteFolder fld, true
            If Err.Number <> 0 Then
               msg = "(Error: "&iif( Err.Description = "", "?", Err.Description)&")"
               status = false
            Else
               status = true
            End If
         Else
            msg = "(Folder not found!)"
         End If
         
         ''' Validate '''
         If status = false Then 
            msgLog "DeleteFolder(): Folder was not deleted. "&msg
         Else
            msgLog "DeleteFolder(): Folder deleted."
         End If
         msgLog "   [folder] "&fld
         deleteFolder = status
      End Function 
      
      
      '**
      '* Method:       getChasisType
      '* Ver/Author:   1.0.0/RRR
      '* Date:         2006.02.28
      '* Purpose:      Returns ProductType value (Workstation, Server,...)
      '*
      Public Function getChasisType()
         'dim objOS, tmpOS
         
         'For Each objOS in SO.InstancesOf("Win32_SystemEnclosure")
            'getOSType = objOS.
         'Next
         getChasisType = ""'tmpOS
      End Function

      
      '**
      '* Method:       getOSVersion
      '* Ver/Author:   1.0.0/RRR
      '* Date:         2003
      '* Purpose:      Returns OS version numeric value only (i.e., 5.1.xxx.x)
      '* Rev-Info:     2006.11.29/RRR, replace return value ~NULL~ to null
      '*
      Public Function getOSVersion( cpuname) 
         On Error Resume Next
         dim objOS, ver, rmt
         
         set rmt  = GetObject( "winmgmts://"&cpuname&"/")
         If Err.Number <> 0 Then
      	   msgLog "GetOSVersion(): Unable to access WMI object on "&cpuname&"."
      	   getOSVersion = null
      	   Exit Function
         Else
            for each objOS in rmt.InstancesOf("Win32_OperatingSystem")
               getOSVersion = objOS.Version
            next
            msgLog "GetOSVersion(): "&tmpOS
         End If    
      End Function
      
      
      '**
      '* Method:       getSystemByAlias
      '* Ver/Author:   1.0.0/RRR
      '* Date:         2006.07.25
      '* Purpose:      Returns current system name
      '*
      Public Function getSystemByAlias( alias)
         On Error Resume Next
         dim rmts, cpu, cpuName : Set rmts = GetObject("winmgmts:\\" _
             & alias & "\root\CIMV2").ExecQuery( _
             "SELECT * FROM Win32_ComputerSystem",,48) 
             
         msgLog "GetSystemByAlias(): Retrieve cname for alias """&alias&""""      
         for Each cpu in rmts 
             cpuName = cpu.Name
         next
         
         ''' Validate '''
         If isEmpty( cpuName) Or _
            cpuName = "" Then
            cpuName = null
            msgLog "   [Unresolved]"
         Else
            msgLog "   ["&cpuName&"]"
         End If
         
         set rmts = Nothing
         set cpu  = Nothing
         getSystemByAlias = UCase( cpuName)
      End Function
      
      
      '**
      '* Method:       regDelete
      '* Ver/Author:   1.0.0/RRR
      '* Date:         2006.05.03
      '* Purpose:      Delete File
      '*
      Public Function regDelete( key)
         On Error Resume Next
         'TODO: Test this function
   
         dim status : status = false
   
         WSO.RegDelete key
         If Err.Number <> 0 Then
            msgLog "RegDelete(): Key was not deleted. "
            msgLog "!"&key
            msgLog "!Error: "&iif( Err.Description = "", "?", Err.Description)
            status = false
         Else
            msgLog "RegDelete(): Key deleted. "
            msgLog "   ["&key&"]"
            status = true
         End If
         regDelete = status
      End Function 
      

      '**
      '* Method:       regRead
      '* Ver/Author:   1.0.0/RRR
      '* Date:         2006.07.03
      '* Purpose:      Read key
      '*
      Public Function regRead( key)
         On Error Resume Next
         
         dim val : val = WSO.RegRead( key)
         If Err.Number <> 0 Then
            msgLog "RegRead(): Key was not found."
            msgLog "!"&key
            msgLog "!Error: "&iif( Err.Description = "", "?", Err.Description)
            val = null
         Else
            msgLog "RegRead(): "&key
            msgLog "   [value] "&val
         End If
         
            regRead = val
      End Function 
   
         
      '**
      '* Method:       regWrite
      '* Ver/Author:   1.0.1/RRR
      '* Date:         2006.03.29
      '* Purpose:      Write key
      '*
      Public Function regWrite( key, val, t)
         On Error Resume Next
         'TODO: Test this function
   
         dim status : status = false
                
         WSO.RegWrite key, val, t
         If Err.Number <> 0 Then
            msgLog "RegWrite(): Invalid key ["&key&"]"
            msgLog "!Error: "&iif( Err.Description = "", "?", Err.Description)
         Else
            msgLog "RegWrite(): Write Value["&val&"] to Key["&key&"] (Type["&t&"])" 
            status = true
         End If
         regWrite = status
      End Function
   
   
      '**
      '* Method:       runCommand
      '* Ver/Author:   1.0.0/RRR
      '* Date:         2006.06.06
      '* Purpose:      Run shell command
      '*
      Public Function runCommand( cmdLine)
         On Error Resume Next
   
         msgLog "RunCommand(): "& cmdLine
         dim rCode : rCode = WSO.Run( cmdLine, 0, true)
         If Err.Number <> 0 Then
            msgLog "!Error occured: #"&Err.Number
            msgLog "!Description: "&iif( Err.Description = "", "?", Err.Description)
            rCode = -1 'Invalid Command
         Else
            dim exitMsg : exitMsg = "!Command exited with: "& rCode
            If rCode = 0 Then exitMsg = exitMsg & " (successful)"
            msgLog exitMsg
         End If
         runCommand = rCode
      End Function


      '**
      '* Method:  Path_
      '* Author:  RRR
      '* Date:    2012.02.03
      '* Purpose: (same Path() without logging).
      Function runCommand_( cmdLine)
         logger "#off"
         runCommand_ = runCommand( cmdLine)
         logger "#on"
      End Function
   
   
      '**
      '* Method:       service
      '* Ver/Author:   1.1.0/RRR
      '* Date:         2006.03.30
      '* Purpose:      Service manipulator.
      '*
      '* Rev-Info:     06.06.06/RRR, Method renamed
      '*
      Public Function service( act, srv)
         dim propCollection : Set propCollection = SO.InstancesOf( "Win32_Service")
         dim propObject, serviceInst : serviceInst = False
         dim status : status = ERRO
         
         For each propObject in propCollection
            If propObject.Name = srv Then
               msgLog "Service(): """&srv&""" service is installed."
               serviceInst = True

               Select case LCase( act)
                  case "","status"
                     If propObject.Started = true Then 
                        msgLog "!Service is running."
                        service = L_SRV_RUNNING
                        Exit Function
                     End If
                     If propObject.Started <> true Then 
                        msgLog "!Service is stopped."
                        service = L_SRV_STOPPED
                        Exit Function
                     End If
                     If propObject.StartMode = "Automatic" Or propObject.StartMode = "Auto" then 
                        msgLog "!Service start mode is set to ""Automatic""."
                        service = L_SRV_AUTO
                        Exit Function
                     End If
                     If propObject.StartMode = "Manual" Then 
                        msgLog "!Service start mode is set to ""Manual""."
                        service = L_SRV_MANUAL
                        Exit Function
                     End If
                     If propObject.StartMode = "Disabled" then 
                        msgLog "!Service is disabled."
                        service = L_SRV_DISABLED
                        Exit Function
                     End If
                  case "start"
                     status = serviceStart( propObject)
                  case "stop"
                     status = serviceStop( propObject)
                  case "disabled"
                     status = serviceMode( propObject, "Disabled")
                  case "auto"
                     status = serviceMode( propObject, "Automatic")
                  case "manual"
                     status = serviceMode( propObject, "Manual")
               End Select
            End If
         Next
         If serviceInst = False Then
            msgLog "Service(): """&srv&""" service is not installed."
            status = INVALID
         End If
   
         service = status
      End Function
   

      '**
      '* Method:       serviceMode
      '* Ver/Author:   1.1.0/RRR
      '* Date:         2006.03.30
      '* Purpose:      Set service start mode to: Automatic | Manual | Disabled
      '*
      Private Function serviceMode( srv, mode)
         dim iter : iter = 0
         dim modeType : modeType = 0
         dim status : status = ERRO
   
         If mode = "Disabled" Then 
            If srv.Started <> false Then serviceStop( srv)
         End If

         Select case mode
            case "Automatic"
               modeType = L_SRV_AUTO
            case "Manual"
               modeType = L_SRV_MANUAL
            case "Disabled"
               modeType = L_SRV_DISABLED
         End Select
   
         If srv.StartMode = mode Then 
            msgLog "!Service start mode is already set to """&mode&"""."
            serviceMode = modeType
            Exit Function
         Else
            srv.ChangeStartMode( mode) 
            msgLog "!Set service start mode to """&mode&"""."
         End If
   
         ''' Validate '''
         If service( mode, srv.Name) <> modeType Then 
            msgLog "!Failed to set service start mode."
            status = ERRO
         Else
            msgLog "!Service start mode is set to "&mode&"."
            status = modeType
         End If
         serviceMode = status
      End Function
   

      '**
      '* Method:       serviceStart
      '* Ver/Author:   1.1.0/RRR
      '* Date:         2006.03.30
      '* Purpose:      Start a service
      '*
      Private Function serviceStart( srv)
         dim iter : iter = 0
         dim status : status = ERRO
   
         If srv.StartMode = "Disabled" Then
            serviceMode srv, "Manual"
         End If

         msgLog "!Starting service."

         If srv.Started = true Then
            msgLog "!Service is already Started."
            serviceStart = L_SRV_RUNNING
            Exit Function
         Else
            srv.StartService
            do While srv.Started <> true
                WScript.Sleep( L_STD_WAIT)
               iter = iter + 1
               If service( "status", srv.Name) = L_SRV_RUNNING Then Exit Do
               If iter >= 6 Then
                  msgLog "!Timeout while starting service. ("&((L_STD_WAIT/1000)*iter)&" sec.)"
                  Exit Do
               End If
            Loop
         End If
   
         ''' Validate '''
         If service( "status", srv.Name) <> L_SRV_RUNNING Then 
            msgLog "!Failed to start service."
            status = ERRO
         Else
            msgLog "!Service has been started."
            status = L_SRV_RUNNING
         End If
         serviceStart = status
      End Function
   

      '**
      '* Method:       serviceStop
      '* Ver/Author:   1.1.0/RRR
      '* Date:         2006.03.30
      '* Purpose:      Stop a service
      '*
      Private Function serviceStop( srv)
         dim iter : iter = 0
         dim status : status = ERRO
   
         If srv.StartMode = "Disabled" Then
            msgLog "!Service is disabled."
            serviceStop = L_SRV_DISABLED
            Exit Function
         End If
   
         msgLog "!Stopping service."

         If srv.Started <> true Then
            msgLog "!Service is already Stopped."
            serviceStop = L_SRV_STOPPED
            Exit Function
         Else
            srv.StopService
            do While srv.Started <> false
               WScript.Sleep( L_STD_WAIT)
               iter = iter + 1
               If service( "status", srv.Name) = L_SRV_STOPPED Then Exit Do
               If iter >= 6 Then 
                  msgLog "!Timeout while stopping service. ("&((L_STD_WAIT/1000)*iter)&" sec.)"
                  Exit Do
               End If
            Loop
         End If

         ''' Validate '''
         If service( "status", srv.Name) <> L_SRV_STOPPED Then 
            msgLog "!Failed to stop service."
            status = ERRO
         Else
            msgLog "!Service has been stopped."
            status = L_SRV_STOPPED
         End If
         serviceStop = status
      End Function
   End Class   

   
''' HISTORY '''

' 20060724 LOECmd  1.0.0.0  Ported methods and properties from common-loe-1.0.2.wsc:
'                              copyFile()
'                              copyFolder()
'                              deleteFile()
'                              deleteFolder()
'                              getChasisType()
'                              getOSVersion()
'                              regDelete()
'                              regRead()
'                              regWrite()
'                              runCommand()
'                              service()
'                              serviceMode()
'                              serviceStart()
'                              serviceStop()
'                           Convert content into Class (LOECmd)
'                           Implemented timeout in service methods
' 20060725 LOECmd  1.0.0.2  Implemented getSystemByAlias(), to retrieve BIOS name from dns alias
' 20060731 LOECmd  1.0.0.3  Optimized deleteFile() to avoid 'Permission Access Denied' messages
' 20061129 LOECmd  1.0.0.4  Optimized getOSVersion() to return null value
'                           Optimized copyFile() and copyFolder() to not do anything if src = trg
' 20061207 LOECmd  2.0.0.0  Merge libutil-1.0.1 with libloe-1.0.1
' 20070404 LOECmd  1.0.2.5  Unmerge from libutil-2.0.0-vbs
' 20070404 LOECmd  1.0.2.6  Fixed bug with service Timeout mechanism
' 20070503 SYSCmd  1.0.2.7  Rename libloe/LOECmd to libscmd/SYSCmd
' 20070620 SYSCmd  1.0.2.8  Optimitized formatting on all Err.Description clauses
' 20090729 SYSCmd  1.0.2.9  Fixed bugs: with service status and timeout routine, and remove 
'                                                   timeout in service mode and improve validation
' 20120203 SYSCmd  1.0.2.10 Added runCommand_()
' TODO                      getChasisType()
' TODO                      RmtProcess class to send command line to remote systems
' TODO                      Add Rename() function
