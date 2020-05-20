'
' For usage instructions see: https://github.com/ascii78/pdf-handler
'
' Make sure to set PDFexe and PdfDir below
Dim PdfExe
' https://kb.foxitsoftware.com/hc/en-us/articles/360042671711-Parameters-for-Opening-PDF-Files-with-a-command
pdfExe="""C:\Program Files (x86)\Foxit Software\Foxit Reader\FoxitReader.exe"""
' https://help.tracker-software.com/pdfxe8/index.html?command-line-options_ed.html
' pdfExe="""C:\Program Files\Tracker Software\PDF Editor\PDFXEdit.exe"""
Dim PdfDir

' change this:
PdfDir="C:\Users\username\Documents\pdf\"

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
