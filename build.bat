mkdir bin
del /Q .\bin\*
copy .\asm\* .\bin
c:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /out:.\bin\gerip.exe gerip.cs /r:".\asm\WatiN.Core.dll"
