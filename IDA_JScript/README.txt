

dependancies:
   spSubclass.dll - activex must regsvr32, in dependancies.zip
   SCIVBX.ocx     - activex must regsvr32, in this dir
   msscript.ocx   - activex must regsvr32, probably already have
   vb6 runtimes   - from MS probably already installed
   SciLexer.dll   - must be in same dir as SCIVBX.ocx, in this dir
   IDASrvr.plw    - must install this IDA plugin, see ./../IDASrvr/bin

This is a standalone interface to interact and script commands sent to IDA
through the IDASrvr plugin using Javascript.

This interface uses the scintinella control which provides syntax highlighting,
intellisense, and tool tip prototypes for the IDA api which it provides. It has been deisgned
as an out of process UI for ease of development and so more complex features
could be added.

Should support most of the commonly used api. If you need to get fancy its easy
to add more features using the template.

Note this will only connect to the last instance of IDA open at present. Support
for multiple clients, and ferrying data between them is possible, it just takes
some minor changes to how IPC server handles are looked up. (now uses static server name
in plugin)

For the ida function list see file api.api it has all the prototypes.
The main class to access these functions is "ida." 

The intellisense is kind of basic, it doesnt recgonize which object is being
accessed?! so "anything." will falsly bring up the same list. not sure why this is.
I didnt code it, but its useful even with this bug.

Also there are a couple wrapped functions available by default without a class
prefix. 

h(x) convert x to hex //no error handling in this yet..also high numbers can overflow error (dll addr)
alert(x) supports arrays and other types
t(x) appends x to the output textbox on main form.


