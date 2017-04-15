# CL Converter OpenXml

Simple VB.Net program to process specific archives of rtf files by de-archiving, converting to docx format, 
renaming according to specific filename scheme (hard-coded) and re-zipping converted files.

For the moment, working in two branches, master and ooxml.
This is OpenXml (ooxml) branch, using Microsoft's OpenXml SDK to create docx's and a RichTextbox control to read and extract text
from rtf files. Much faster. Still have problems with table conversion, which the RichTextbox control handles unsatisfactorily.

# Installation

Download zip archive, unzip and run setup.exe. Clickonce deployment used, so no administrator privileges required. The application will self-update upon launch. It looks for new version from local network, in a folder on Q drive (GSC ressouce)
