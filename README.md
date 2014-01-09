LVA
===

Better Structured VBScript by loading other VBScript files as your API.

******************************************************************************
<pre>
OPTION EXPLICIT

   Execute library( "lib-utility-1.0.4.vbs")
   
   '---BEGIN--------------------------------------
   logger left(getScriptName(),len(getScriptName())-4)&".log" : msgLog "" '-- blank line
   msgLog "Running... ("&getScriptName()&")"
   
   '<YOUR SCRIPT GOES HERE>
   
   '-- fini
   ExitProcess Null, 0

   '---END----------------------------------------


   '-- Import library into memory (2006.07.21/1.0.0/RRR)
   Function library( ByVal libname)
      dim libf:set libf=CreateObject("Scripting.FileSystemObject"):if not libf.fileExists(libname) then wscript.stdOut.write " Error: Could not locate library: "&libname:WScript.Quit( 1)
      library=libf.OpenTextFile(libname,1).ReadAll()
   End Function
</pre>
