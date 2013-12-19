
This is a quick Wingraph32/qwingraph replacement that has some extra features
such as being able to navigate IDA to the selected nodes when they are clicked
on in graph view, as well as being able to rename the selected node from the 
graph, or adding a prefix to all child nodes below it.

This application requires the .NET runtime 3.5 or greater, as well as the IDASrvr
plugin (project found in the parent directory) to integrate with IDA.

This uses the GLEE graphing library from Microsoft Research which is free for 
non-commercial use.

To install, rename the default wingraph32 or qwingraph executable from the IDA
home directory, and replace it with the one provided.  