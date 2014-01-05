
Flash Patcher v1.0
-------------------------
ui Author: David Zimmer <dzzie@yahoo.com>

swfdump   copyright Adobe
hexed.ocx copyright Rang3r
zlib.dll  copyright Jean-loup Gailly and Mark Adler
cmdoutput.bas copyright Joacim Andersson, Brixoft Software


install requirements:
--------------------------

if pdfstreamdumper is installed you can skip this section
except that

--> swfdump requires the java runtime installed.

hexed.ocx - drag and drop on reg_ocx.bat to register 

vb6 runtimes (probably already installed)
http://www.microsoft.com/en-us/download/details.aspx?id=24417

mscomctl.ocx  (probably already installed)  
http://www.microsoft.com/en-us/download/details.aspx?id=10019


couple notes:
--------------------------
when run against a target swf, it will
create a decompressed version of the swf and a .txt disasm log file
these files will be cached and used on subsequent loads. if you wish to
start over from scratch use the tools->delete cached * options.

any nops or other edits you do to the swf through the built in hexeditor
are done to the decompressed swf only, never the original. that way if you
mess up you can just use the delete cached option and start over. 

this app is really just a parser/viewer/hexeditor built around the 
Adobe swfdump utility. that makes it cheesy and brilliant all at the same time :)

this app was thrown together in a weekend with no prior knowledge of
the internals of flash opcodes or disasm. 


htmlview notes:
---------------------------

this was a real quick feature i added for usability.

click on a label to rename it globally (use a unique name or you could stomp on other names)
click on a var_* to rename it
you can add comments at the end of any line of disasm by just clicking and typing.

