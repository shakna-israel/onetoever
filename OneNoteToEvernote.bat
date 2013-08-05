REM Set up folders.
IF exist C:\ExportFolder ( echo C:\ExportFolder exists ) ELSE ( mkdir C:\ExportFolder && echo C:\ExportFolder created)
REM Export from Onenote
START /WAIT OneNoteExport.vbs
REM Import to Evernote
for /f %i in ('dir /b C:\ExportFolder\') do "C:\Program Files (x86)\Evernote\Evernote\Evernote.exe" %i
REM Sync Evernote
"C:\Program Files (x86)\Evernote\Evernote\Evernote.exe" /Task:SyncDatabase