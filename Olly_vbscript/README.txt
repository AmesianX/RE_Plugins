

Author: dzzie@yahoo.com

Note: this plugin also includes IDA_Sync feature which opens
a command socket and can send commands to my IDAVBScript plugin
so you can view the disasm in IDA in sync with your debugging

Remeber to register Olly_vbScript.dll if you dont compile it from
vb source. the other dll is the olly plugin and has to be in your
plugin directory.


What is it?
-----------------------------------------------------
This is an OllyDbg plugin that gives you a scriptable
interface to the native OllyDbg plugin API.

With it you can analyze and patch code. Set breakpoints,
control execution, and respond to breakpoint events in 
an automated fashion from the comfort of a HLL language 
interface using standard VBscript.

See the file protos.txt for a list of extra functions that
this thing supports.

Want to add more? create a COM object can be used with the 
windows script host and do a CreateObject()



Installation:
------------------------------------------------------

copy all of the dlls and mdb database to your olly plugins directory

(defined under options -> appearance - > directories)

from a command line run the following commands:

regsvr32 Olly_vbScript.dll

MS dependancies: msscript.ocx,msado15.dll,mscomctl.ocx


Howto:
-------------------------------------------------------

You can have your scripts work across breakpoints.

Calling any of the execution functions such as go will
relinquish control back to olly. If you have set a breakpoint
AND _enabled_bpxhooking_ then you have the script reinvoked at 
each breakpoint.

There are 3 functions it will look for when a breakpoint is
hit. The first mechanism goes by breakpoint index. like
1st time we hit a breakpoint 2nd time etc...these functions
are named in the convention:

sub BpxHandler_1
sub BpxHandler_2
..etc...

You can check to see what is the current step and reset it
with the functions:

Property Get BpxHandler()
sub SetBpxHandler(x)

The second mechanism calls breakpoint handlers based on the
current instruction pointer. Subs defined to catch these breakpoints
will be in the format 

sub BpxHandler_?

Where ? is the hex(eip) value of the bpx

The last type of handler you can use to catch breakpoint events is a
generic handler (like if you had to single step allot) This function
has to be named:

sub Default_BPX_Handler

The search order for bpxhandlers is

1) by eip
2) by index
3) default handler

If none of these are found then you will get an alertbox stating what it
couldnt find.

Rember to use this feature you have to enable and disable it with the 
functions 

sub EnableBpxHook()
sub DisableBpxHook()


Sample script:
---------------------------------------------------------------------
EnableBpxHook
nextCmd = eip+instlen(eip)
setbpx ( nextCmd  )  
refresh
go

sub BPXHandler_1
   nop nextCmd , 10
   refresh
   disablebpxhook
   showui
end sub


ToDo:
----------------------------------------------------------------------

I have a full syntax highlight IDE complete with intellisense i might
integrate with this plugin...As much as I love intellisense and function
prototypes...I am debating weather i want to really add the bloat among other
considerations.

In terms of scripting functionality...This build has 95% of everythign useful
I think....mabey a couple more macro type access things and a couple more
things will come up...but it seems like a comfortable level of access for now.

Other function prototype i am thinking of include:

LibraryBpx(lib.fxname)
SetCaption(newText)
Hidedebugger(optional bpxonAccess = false) 'mod PEB, optional hwrbpx on read access to field
ListModules
ListModFunctions

