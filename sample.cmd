@echo off

setlocal
 
  set __wd__=%~dp0%
  set __wd__=%__wd__:~,-1%

  :: run in current directory
  pushd %__wd__%
 
  cscript /nologo sample.vbs


  popd

endlocal

echo on
