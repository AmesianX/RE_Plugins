
Handfull of misc olly and ida plugins I have created

IDA_Jscript:
   Script IDA in Javascript. uses IDASrvr to listen to WM_COPYDATA messages.
   script interface is in external process for ease of development. Uses scintinella
   control for syntax highlighting, and intellisense features. See screen shot.   

olly_hittrace: 
   You set breakpoints in the UI and it will then run
   the app automating it and logging which ones were hit.

Olly_module_bpx:
   allow you to set breakpoints within modules which have not yet been loaded.

Olly_vbscript:
   vbscript automation capability for olly including working across breakpoint events.
   also includes ollysync which is udp based remote control to query and control olly
   also can be used to send display sync commands to idavbscript while stepping in olly

OllySrvr:
   wm_copydata IPC server running in olly. allows you to query data from current process
   and sync display. 

IDASrvr: 
   wm_copydata IPC server running in IDA. allows you to send commands to IDA from another
   process to query data and control interface display. used from my uGrapher wingraph32 
   replacement so when you click on a node it will navigate to the function in IDA.
   WM_COPYDATA was chosen because it is synchronous

uGrapher:
   replaces default wingraph32, uses uDrawGraph. graphing up, still needs features. once
   idasrvr working again, it can navigate to nodes as you click then in the graph..not sure
   i like uDraw graph after having used some other activeX controls. I already have this 
   working and more in another way using the socket based ida control in idavbscript. but 
   the uDrawGraph classes are nice and debugged. ported from pedrams python work on openrce.

IdaVbScript:
   ton of small tools for IDA all thrown into one interface. see screenshot.png for details.
   also include a udp based command socket to send commands and receive data from IDA from 
   external apps. Used by olly_vbscript and wingraph32 replacement. (not uGrapher)

Wingraph32:
   another wingraph32 replacement using a 3rd party graphing ocx (see readme), see screen 
   shot for features. can also navigate IDA to selected nodes using the UDP socket in 
   idavbscript plugin.

gleegraph:
   c# wingraph32 replacement using Microsoft GLEE library (free for non-commercial use)
   click on a node to navigate IDA to it. Right click menu to also rename functions within
   IDA, or prefix all nodes below. Uses IDASrvr see screen shot.