call "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\Tools\VsDevCmd.bat"
cd Properties
resgen.exe /compile Resources.de.resx,MeetingInfo.Resources.de.resources
Al.exe /t:lib /embed:MeetingInfo.Resources.de.resources /culture:"de" /out:"MeetingInfo.resources.dll"
del MeetingInfo.Resources.de.resources
cd ".."
if not exist "bin/de/" mkdir "bin/de/"
move "Properties\MeetingInfo.resources.dll" "bin/de/"
pause