# clconverter

CL Converter

Simple vb.net program to process sepcific archives of rtf files by de-archiving, converting to docx format, 
renaming according to specific filename scheme (hard-coded) and re-zipping converted files.

For the moment, working in two branches, master and ooxml. Master is using a Word application to handle rtf conversion
to docx. Slow. Ooxml is using Microsoft's OpenXml SDK to create docxs and a RichTextbox control to read and extract text
from rtf files. Much faster. Still have problems with table conversion, which the RichTextbox control handles unsatisfactorily.
