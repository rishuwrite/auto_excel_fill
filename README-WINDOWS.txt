FBT Manifest Generator - Windows

QUICK START
1. Extract FBT-Manifest-Generator-Windows.zip to a normal folder you can write to.
2. Double-click FBT-Manifest-Generator.exe.
3. Select your FedEx booking .xls export.
4. Click GENERATE FBT FILES.
5. The program creates the US and/or Non-US FBT .xls files in the selected folder.

ADMIN RIGHTS
Normally none are required if Microsoft Excel is already installed on the workstation.
The program uses installed Excel automatically on Windows, which avoids installing
LibreOffice or Python.

If Excel is not installed, the command-line version can use LibreOffice instead.
On a locked corporate workstation, IT/security policy may still block unsigned EXE
files even when administrator rights are not required. That is a Windows policy issue,
not a requirement of the generator.

DATA
The GUI does not upload the booking file to GitHub or any web service. Processing is
local on the workstation.

CURRENT TEMPLATE
The bundled GUI is configured for the supplied NFEI YES templates. If NFEI NO templates
are later provided, the GUI can be extended to offer that mode too.
