@echo off

::Starting WinSCP and preparing respective folders by deleting outdated data
echo Starting WinSCP
del C:\Temp\SY_Data\Ngeri\*.* /q
del C:\Temp\SY_Data\Ogongo\*.* /q
del C:\Temp\SY_Data\Humanist\*.* /q
del C:\Temp\SY_Data\Sena\*.* /q
del C:\Temp\SY_Data\Sibuoche\*.* /q
del C:\Temp\SY_Data\Midida\*.* /q
del C:\Temp\SY_Data\Oyani\*.* /q
del C:\Temp\SY_Data\Kisegi\*.* /q
del C:\Temp\SY_Data\Ongo\*.* /q
del C:\Temp\SY_Data\Nyatoto\*.* /q
del C:\Temp\SY_Data\Kitare\*.* /q

::Calling WinSCP code file named winscp-scriptsy.text to download respective databases from remote ftp site 
"C:\Program Files (x86)\WinSCP\WinSCP.com" /script="C:\Temp\SY_Data\winscp-scriptsy.txt"
"C:\Program Files\7-Zip\7z.exe" e C:\Temp\SY_Data\Ngeri\*.7z -oC:\Temp\SY_Data\Ngeri
"C:\Program Files\7-Zip\7z.exe" e C:\Temp\SY_Data\Ogongo\*.7z -oC:\Temp\SY_Data\Ogongo
"C:\Program Files\7-Zip\7z.exe" e C:\Temp\SY_Data\Humanist\*.7z -oC:\Temp\SY_Data\Humanist
"C:\Program Files\7-Zip\7z.exe" e C:\Temp\SY_Data\Sena\*.7z -oC:\Temp\SY_Data\Sena
"C:\Program Files\7-Zip\7z.exe" e C:\Temp\SY_Data\Sibuoche\*.7z -oC:\Temp\SY_Data\Sibuoche
"C:\Program Files\7-Zip\7z.exe" e C:\Temp\SY_Data\Midida\*.7z -oC:\Temp\SY_Data\Midida
"C:\Program Files\7-Zip\7z.exe" e C:\Temp\SY_Data\Oyani\*.7z -oC:\Temp\SY_Data\Oyani
"C:\Program Files\7-Zip\7z.exe" e C:\Temp\SY_Data\Kisegi\*.7z -oC:\Temp\SY_Data\Kisegi
"C:\Program Files\7-Zip\7z.exe" e C:\Temp\SY_Data\Ongo\*.7z -oC:\Temp\SY_Data\Ongo
"C:\Program Files\7-Zip\7z.exe" e C:\Temp\SY_Data\Nyatoto\*.7z -oC:\Temp\SY_Data\Nyatoto
"C:\Program Files\7-Zip\7z.exe" e C:\Temp\SY_Data\Kitare\*.7z -oC:\Temp\SY_Data\Kitare

echo WinSCP finished


