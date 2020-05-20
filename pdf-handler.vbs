'
' pdf url handler, this accepts URLs in the form of pdfshortname + pagenum like:
'
' pdf://shortfilename:40
'
' note that:
'   - there is NO pdf extension needed for filename, this is to keep the url short
'   - please change PdfDir in script to where your pdfs are located
'   - advise: if you need to type the url a lot, try renaming your file to 
'     something short like: 'Very Long Book Title.pdf', 'vlbt.pdf'
'   - choose a pdf reader below, these are the only ones that work by actually changing
'     the page in the _current_ tab. almost all others will open a new tab or
'     window even if the file is already open (eg. acrobatreader)
'   - for this to work you need an entry in the registry
'
' if you are a developer:
'   - i've never used vbs before but:
'       - python: could not let me open files in the same tab by using Popen
'       - powershell: shows/flashes a short cmd/console window which you can
'                     not get rid off.
'       - vbscript: works and is installed by default on win10
'
' DISCLAIMER:
'
'   - if someone sends you a pdf:// link and bad things happen, then welcome
'     to the nineties...
'
' copy the following for to a .reg file, remove leading ', point the command 
' location to where you copied this file.

' ---

' Windows Registry Editor Version 5.00

' [HKEY_CLASSES_ROOT\pdf]
' @="URL:pdf Protocol"
' "URL Protocol"=""

' [HKEY_CLASSES_ROOT\pdf\shell]

' [HKEY_CLASSES_ROOT\pdf\shell\open]

' [HKEY_CLASSES_ROOT\pdf\shell\open\command]
' @="wscript \"C:\\Users\\username\\Documents\\pdf-handler.vbs\"  %1"

'---

' to disable onenote warnings read:
' https://superuser.com/questions/1307645/how-to-disable-hyperlink-security-notice-in-onenote-2016
'

Dim PdfExe
' https://kb.foxitsoftware.com/hc/en-us/articles/360042671711-Parameters-for-Opening-PDF-Files-with-a-command
pdfExe="""C:\Program Files (x86)\Foxit Software\Foxit Reader\FoxitReader.exe"""
' https://help.tracker-software.com/pdfxe8/index.html?command-line-options_ed.html
' pdfExe="""C:\Program Files\Tracker Software\PDF Editor\PDFXEdit.exe"""
Dim PdfDir

' change this:
PdfDir="C:\Users\frido\Documents\pdf\"

Dim RegEx
Set RegEx = New RegExp
RegEx.IgnoreCase = True
RegEx.Global=True
RegEx.Pattern="^pdf://([A-Za-z0-9\_\-]+):([0-9]+)/$"

strURL= WScript.Arguments.Item(0)

If NOT RegEx.Test(strURL) Then
    WScript.Echo strURL & " does NOT look like a valid PDF URL. Try: " & RegEx.Pattern
    WScript.Quit 1
Else
    Set colMatches=RegEx.Execute(strURL)
    For Each match In colMatches
        if match.Submatches.Count = 2 Then
            file=match.Submatches.Item(0)
            page=match.Submatches.Item(1)
            Set objShell = CreateObject("Wscript.Shell")
            objShell.Run(pdfExe & " /A page=" & page & " " & pdfdir & file & ".pdf")
        Else
            WScript.echo "Matchcount wrong: " & match.Submatches.count
        End if
    Next    
End If