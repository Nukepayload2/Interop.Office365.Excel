call "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\Tools\VsDevCmd.bat"

tlbimp "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\OFFICE16\mso.dll" /out:"%~dp0\Microsoft.Office.Interop.Core.dll"
tlbimp "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB" /reference:"%~dp0\Microsoft.Office.Interop.Core.dll" /out:"%~dp0\Microsoft.Vbe.Interop.dll"
tlbimp "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" /reference:"%~dp0\Microsoft.Office.Interop.Core.dll" /reference:"%~dp0\Microsoft.Vbe.Interop.dll" /out:"%~dp0\Microsoft.Office.Interop.Excel.dll"
pause