@echo off
@rem Console driver for the VM.
copy /y nanovm.js + vmc.js _vmc.js
cscript /nologo _vmc.js %*
