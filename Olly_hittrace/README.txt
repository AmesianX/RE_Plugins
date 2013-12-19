
Updated 10-22-2010
   x- ascii dump feature   
   x- support multiple trace expressions per breakpoint
   x- grab comment from olly disasm to add to trace message
   x- log output to file
   x- added a hexdump feature
   x- ability to edit a saved breakpoint (must re-add when done)
   x- show hitcount stats (seperete report use Show HitCount button)

HitTrace - david zimmer <dzzie@yahoo.com>

This is a simple plugin based on my modulebpx code.

You set breakpoints in the UI and it will then run
the app automating it and logging which ones were hit.

To set breakpoints in the main module use "main"
(no quotes) as mod name

For dlls you can enter partial strings as long
as they are unique and as long as they are the 
dlls actual name as found in its export table.

Addresses of breakpoints are set in rva's (in case
a dll gets rebased)

The LogExp is optional..it accepts any type of
expression the ollydbg expression window takes
such as [ebp+4] or eax or whatever. If this is 
set and is valid then it will shoot the results to the 
log window on breakpoint. Sorry only supports one
expression to evaluate per bpx right now.

Each breakpoint is assigned an index which is visible
in the listbox. You can use one of these indexes
in the abort box to have it bail on tracing when
that bp is reached.

For example you can run it on looper.exe with the 
following settings

main	1030	[esp+4]
main	1070	[esp+4]
main	1136	

abort on index: 2

Then hit View results to see hit count and log window

or load the looper.htl sample provided (abort on index is not saved to htl files)

if you recompile looper offsets may change disasm to find new ones.

added: 10-22-2010 - 

if the first letter of your expression is A, then it will take the 
value of the expression and consider it a string offset and try to do a string dump of
that address.

if the first letter of your expression is H, then it will take the 
value of the expression and consider it a memory offset and try to do a hex dump of
that address. you specify the lengh to dump right after the H like H8 dumps 8 chars.

use str_looper.exe and str_looper.htl to demo this feature.

strlooper demo

.text:004010F3                 push    eax        ;encoded string address
.text:004010F4                 call    j_Decoder
.text:004010F9                 add     esp, 0Ch   ;eax = decoded string 

so you could use
   main  10f3  eax      to get the address
   main  10f9  a eax    to get an ascii dump of the decoded string
   main  10f9  h8 eax   to get an 8 char hexdump of the decoded data




