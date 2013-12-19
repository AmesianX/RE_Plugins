

This is a small (and old!) plugin used to control Olly through
a WM_COPYDATA interprocess communication channel. 

This method was chosen over sockets because of its synchronous manner.

Clients will have to set thier "IPCNAME" under the following regkey

Software\\VB and VBA Program Settings\\IPC\\Handles

This plugin will create a key "OLLY_SERVER" with the default value 
being the window handle its using to listen to messages on.

Clients specify their window handle in the registry as well, and pass
in their regkey IPC name in the wm_copyData message they submit.

I cant really remember why i did it this way entirly..probably because
this is so old!

commands this accepts are:

/* OLLY_SERVER Command format
	msg:message
	setbp:lngadr
	killbp:lngadr
	getbase:IPC_CLIENTNAME  (returns va)
	setname:address:name
	setcomment:address:comment
	getname:address:IPCRECV
	getcomment:address:IPCRECV
	getasm:offset:IPCNAME
	setbp_modrva:module:rva
	killbp_modrva:module:rva
*/



