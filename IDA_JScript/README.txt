

dependancies:
   spSubclass.dll - activex 
   SCIVBX.ocx     - activex 
   msscript.ocx   - activex  
   vb6 runtimes   - from MS probably already have installed
   SciLexer.dll   - must be in same dir as SCIVBX.ocx, in this dir
   IDASrvr.plw    - must install this IDA plugin, see ./../IDASrvr/bin

To register the activeX dependancies included in this download, double click
the install.bat script. 

This is a standalone interface to interact and script commands sent to IDA
through the IDASrvr plugin using Javascript.

This interface uses the scintinella control which provides syntax highlighting,
intellisense, and tool tip prototypes for the IDA api which it provides. It has
been deisgned as an out of process UI for ease of development and so more 
complex features could be added.

Should support most of the commonly used api. If you need to get fancy its easy
to add more features using the template.

When IDA_jscript first starts, it will enumerate active IDASrvr instances. If
its only one active it will automatically connect to it, else it will prompt you
to select which one to interact with.

For the ida function list see file api.api it has all the prototypes.
The main class to access these functions is "ida." 

There are a couple wrapped functions available by default without a class
prefix. 

h(x) convert x to hex //no error handling in this yet..also high numbers can overflow error (dll addr)
alert(x) supports arrays and other types
t(x) appends x to the output textbox on main form.


