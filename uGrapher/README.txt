
Dependancies: vbDevKit.dll, spSubclass.dll, default path install of uDrawGraph.exe

'features:
'        loads c:\ida_last_graph.txt or command line file
'        if command line file, save copy as last_graph
'        starts:   C:\Program Files\uDraw(Graph)\bin\uDrawGraph.exe
'        on node selection in uDraw, it can make IDA jump to the function
'        if IDASrvr is installed and running (WM_COPYDATA version)

rename real wingraph32.exe to _wingraph.exe and put this one in its place.
you can still use the original from menu item.

this app is in development and still needs more features, the graphing part
is done though

'This code is based on:
'    uDraw Connector
'    Copyright (C) 2006 Pedram Amini <pedram.amini@gmail.com>
'    contact:      pedram.amini@gmail.com
'    organization: www.openrce.org
'
'Ported to vb by: dzzie@yahoo.com


