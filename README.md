# CL Converter Master (Word)

Simple VB.Net program to process specific archives of rtf files by de-archiving, converting to docx format, 
renaming according to specific filename scheme (hard-coded) and re-zipping converted files.

For the moment, working in two branches, master and ooxml. 
This is master branch, using a Word application to handle rtf conversion to docx. Slow.

Known issues:
- slow. Using a Word app to handle file conversion is oh so slow... spending resources on using huge
app to perform such simple opperations.

To do:
- perhaps try multi-threading if memory sufficient to use 2 word apps in same time ! Crazy?
