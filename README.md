# A short script to open PDF files on a certain page

I use this PDF handler to open PDFs on my local disk directly from OneNote. The only added feature it has above the normal click is that it goes directly to a certain page while retaining the view of the current PDF reader. See below for a demo; both the text and the image have link that will let the PDF reader next to it jump to the correct page without opening new tabs. 

![demo](img/demo.gif)

## Usage

In either select text or image and create a link with <kbd>ctrl+k</kbd> 
or `insert link` button. 

The PDF URLs for calling this script is in the form of shortfilename + pagenum like:

 `pdf://shortfilename:40`

- there is *NO pdf extension* needed for `shortfilename`, this is to keep the URL short.
- advise: if you need to type the URL a lot, try renaming your file to something short like: 'Very Long Book Title.pdf', 'vlbt.pdf'
- if you create a link of a picture, you need to use <kbd>ctrl</kbd>+click, with normal text you can just click.

Below you can see that the URL link in OneNote, the PDFs are on my desktop on the left.

![links](Z:\wrk\40-pdf-handler\img\show-link.jpg)

## Installing

- download .vbs script and put it somewhere on your pc (eg. Documents/)
- change PdfDir in this script to where your pdfs are located
- change PdfExe to either Foxit or PDF-Xchange (others won't work)
- download .reg file 
- change location of command in registry file to where you put the
.vbs file
- execute reg file or add commands manually
- for OneNote optionally disable warnings [warnings](https://superuser.com/questions/1307645/how-to-disable-hyperlink-security-notice-in-onenote-2016)

## For developers

- I've never used vbs before but:
  - python: could not let me open files in the same tab using Popen
  - PowerShell: shows/flashes a short console window which you can not get rid of [#3028](https://github.com/PowerShell/PowerShell/issues/3028)
  - vbscript: works and is installed by default on win10

- regarding pdf readers these are the only ones that work by actually changing the page in the _current_ tab. almost all others will open a
  new tab or window even if the file is already open (eg. acrobat reader)


# DISCLAIMER

 - These pdf:// links are systemwide, if someone sends you a pdf:// link don't click on them.
