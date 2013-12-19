
  Author: David Zimmer <dzzie@yahoo.com>

  Notes: this is ugly and was done as quick as possible but it works
         if you need more than 50 bpx recompile;

		 THIS APP USES DLL NAME LISTED IN MODULES EXPORT TABLE FOR
		 STRING MATCHING. IF DLL NAME IS DIFFERENT IT WILL LOOK LIKE
		 THIS APP ISNT WORKING.

This is a simple plugin designed to allow you to set breakpoints within 
modules which have not yet been loaded.

For module name you can enter partial strings as long as they are unique 
and as long as they are the dlls actual name as found in its export table.

Addresses of breakpoints are set in rva's in case a dll gets rebased  

