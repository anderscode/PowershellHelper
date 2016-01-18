@ECHO OFF
SET ScriptPath=%~dp0
SET /p User1="Username@Domain : "
runas /user:%User1% %ScriptPath%start.bat