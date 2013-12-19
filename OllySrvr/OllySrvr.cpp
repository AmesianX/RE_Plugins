
#include <windows.h>
#include <stdio.h>
#include <StdDef.H> // offsetof()
#include "plugin.h"


typedef struct{
    int dwFlag;
    int cbSize;
    int lpData;
} cpyData;

char *IPC_NAME = "OLLY_SERVER";
HWND ServerHwnd=0;
WNDPROC oldProc=0;
bool m_debug = true;

void strlower(char* x){
	if( (int)x ){
		int y = strlen(x);
		for(int i=0;i<y;i++){
			x[i]=tolower(x[i]);
		}
	}
}

 
//basically ripped from pedrams olly bp man ;)
int FindModuleFromName(char* modName){

	t_table *mod_table;
	t_module *module;
	char buf[250];
	

	mod_table = (t_table *)Plugingetvalue(VAL_MODULES);

	for (int i=0; i<mod_table->data.n; i++){

        if ((module = (t_module *) Getsortedbyselection(&(mod_table->data), i)) == NULL){
            if(m_debug) Addtolist(0, 1, "OllySrvr Unable to resolve module at index %d.", i);
            continue;
        }
		
		if(module->name == NULL) break;
		strncpy(buf, module->name, 249);
		strlower(buf);
		
		if (strstr(modName,"main") != 0 && i==0){
			return module->base;
		}

        if (strstr(buf, modName) != 0){
			 if(m_debug) Addtolist(0, 1, "OllySrvr> Found base for module %s @ %x.", modName, module->base );
		     return module->base;
        }
    }

	if(m_debug) Addtolist(0, 1, "OllySrvr> Module %s not Found", modName);
	return 0;

}


HWND ReadReg(char* name){

	 char baseKey[200] = "Software\\VB and VBA Program Settings\\IPC\\Handles";
	 char tmp[20] = {0};
     unsigned long l = sizeof(tmp);
	 HKEY h;
	 
	 RegOpenKeyEx(HKEY_CURRENT_USER, baseKey, 0, KEY_READ, &h);
	 RegQueryValueExA(h, name, 0,0, (unsigned char*)tmp, &l);
	 RegCloseKey(h);

	 return (HWND)atoi(tmp);
}


void SetReg(char* name, int value){

	 char baseKey[200] = "Software\\VB and VBA Program Settings\\IPC\\Handles";
	 char tmp[20];

	 HKEY ret;
	 sprintf(tmp,"%d",value);
	
	 RegOpenKey(HKEY_CURRENT_USER, baseKey, &ret);
	 RegSetValueEx(ret, name,0, REG_SZ, (const unsigned char*)tmp , strlen(tmp)); 
	 RegCloseKey(ret);

}

bool SendTextMessage(char* name, char *Buffer, int blen) 
{
  
  
  HWND h = ReadReg(name);
  if(IsWindow(h) == 0){
	  Addtolist(0,0,"Could not find valid hwnd for server %s\n",name);
	  return false;
  }

  if(m_debug) Addtolist(0,0,"Trying to send message to %s\n", name);

  cpyData cpStructData;  
  cpStructData.cbSize = blen ;
  cpStructData.lpData = (int)Buffer;
  cpStructData.dwFlag = 3;

  SendMessage(h, WM_COPYDATA, (WPARAM)h,(LPARAM)&cpStructData);  

  return true;

}  

bool SendIntMessage(char* name, int resp){
	char tmp[30]={0};
	sprintf(tmp, "%d", resp);
	if(m_debug) Addtolist(0,0,"SendIntMsg(%s, %s)", name, tmp);
	return SendTextMessage(name,tmp, sizeof(tmp));
}

bool ReceiveTextMessage(int lParam, char *msg, int bufLen){
	
	cpyData CopyData; 
	
	try{
		memcpy((void*)&CopyData, (void*)lParam, sizeof(CopyData));
    
		if( CopyData.dwFlag == 3 ){
			if( CopyData.cbSize >= bufLen ) CopyData.cbSize = bufLen-1;
			memcpy((void*)msg, (void*)CopyData.lpData, CopyData.cbSize);
			return true;    
		}

	}
	catch(...){}

	return false;
}

void HandleMsg(char* m){


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


	const int MAX_ARGS = 6;
	char *args[MAX_ARGS];
	char *token=0;
	char *cmds[] = {"msg","setbp","killbp","getbase","setname", "setcomment", /* 5  */
		            "getname","getcomment","getasm", "setbp_modrva", "killbp_modrva",     /* 10 */
					'\0'  
	               };
	int i=0, x=0;
	int argc=0;
	char buf[500]={0};
	unsigned char tmp[30]={0};

	//split command string into args array
	token = strtok(m,":");
	for(i=0;i<MAX_ARGS;i++){
		args[i] = token;
		token = strtok(NULL,":");
		if(!token) break;
	
	}

	argc=i;

	//get command index from cmds array
	for(i=0; ;i++){
		if( cmds[i][0] == 0){ i = -1; break;}
		if(strcmp(cmds[i],args[0])==0 ) break;
	}

	//handle specific command
	switch(i){
		case -1: Addtolist(0,1,"OllyServer Unknown Command\n");             break; //unknown command
		case  0: Addtolist(0,0,args[1]);			              break; //msg:UI_MESSAGE
		case  1: Setbreakpoint(atoi(args[1]), TY_ACTIVE, 0);	  break; //setbp:lngAdr
		case  2: Setbreakpoint(atoi(args[1]), TY_DISABLED,0);     break; //killbp:lngadr
				  

		case 3:                                          //getbase:IPC_CLIENTNAME  (returns va)
				if(argc<1) Addtolist(0,0,"Invalid arg count to name_va");
				i =  FindModuleFromName("main"); 
				SendIntMessage(args[1], i);
				break;


		case 4: //setname:address:name
				if(argc<2) Addtolist(0,0,"Invalid arg count to setname requires 2");
				Insertname(atoi(args[1]), NM_EXPORT, args[2]) ;
				break;

		case 5: //setcomment:address:comment
				if(argc<2) Addtolist(0,0,"Invalid arg count to setcomment requires 2");
				Insertname(atoi(args[1]), NM_COMMENT, args[2]) ;
				break;

		case 6: //getname:address:IPCRECV
				if(argc<2) Addtolist(0,0,"Invalid arg count to setcomment requires 2");
				if( Findname(atoi(args[1]), NM_ANYNAME, buf) > 0){
					SendTextMessage(args[2], buf, strlen(buf));
				}
				break;

		case 7: //getcomment:address:IPCRECV
				if(argc<2) Addtolist(0,0,"Invalid arg count to setcomment requires 2");
				if( Findname(atoi(args[1]), NM_COMMENT, buf) > 0){
					SendTextMessage(args[2], buf, strlen(buf));
				}
				break;

		case 8: //getasm:offset:IPCNAME
				t_disasm td;
				Readmemory(tmp, atoi(args[1]) , 20, MM_RESTORE);
				i = Disasm(tmp, 20, atoi(args[1]), 0,  &td , DISASM_CODE, Getcputhreadid());
				SendTextMessage(args[2], td.result, strlen(td.result));
				
		case 9: //setbp_modrva:module:rva
		        if(argc<2) Addtolist(0,0,"Invalid arg count to setcomment requires 2");
				i = FindModuleFromName(args[1]);
				x = atoi(args[2]);

				if(i==0){ Addtolist(0,0,"Module not found"); break;}
				if(x > i) {Addtolist(0,0,"Invalid rva used %x", x); break;}

				Setbreakpoint(i+x, TY_ACTIVE, 0);
				break;

		case 10: //killbp_modrva:module:rva
		        if(argc<2) Addtolist(0,0,"Invalid arg count to setcomment requires 2");
				i = FindModuleFromName(args[1]);
				x = atoi(args[2]);

				if(i==0){ Addtolist(0,0,"Module not found"); break;}
				if(x > i) {Addtolist(0,0,"Invalid rva used %x", x); break;}

				Setbreakpoint(i+x, TY_DISABLED, 0);
				break;

				  

	}

	Redrawdisassembler();

};

LRESULT CALLBACK WindowProc(HWND hwnd,UINT uMsg,WPARAM wParam,LPARAM lParam){
	
	char m_msg[2000]={0};
	
	int  i=0;

	if( uMsg = WM_COPYDATA && lParam != 0){
		if( ReceiveTextMessage(lParam, m_msg, 2000) ){
            
			if(m_debug)	Addtolist(0,0,"Message Received: %s \n", m_msg);  

			try{ 
					HandleMsg(m_msg); 
			}
			catch(...){}

	   }
	}

    return 0;
}


void DoEvents() 
{ 
    MSG msg; 
    while (PeekMessage(&msg,0,0,0,PM_NOREMOVE)) { 
        TranslateMessage(&msg); 
        DispatchMessage(&msg); 
    } 

} 

 

BOOL WINAPI DllEntryPoint(HINSTANCE hi,DWORD reason,LPVOID reserved) {
 //  if (reason==DLL_PROCESS_ATTACH) hinst=hi;
  return 1;
};

extc int _export cdecl ODBG_Plugindata(char shortname[32]) {
  strcpy(shortname,"OllyServer - Running");
  return PLUGIN_VERSION;
};

extc int _export cdecl ODBG_Plugininit( int ollydbgversion,HWND hw,ulong *features) {
  
  if (ollydbgversion<PLUGIN_VERSION) return -1;
  
  Addtolist(0,0,"Starting OllyServer - David Zimmer (dzzie@yahoo.com)");

  //immediatly create server window for use (no need to explicitly launch plugin)  
  ServerHwnd = CreateWindow("EDIT","MESSAGE_WINDOW", 0, 0, 0, 0, 0, 0, 0, 0, 0);
  oldProc = (WNDPROC)SetWindowLong(ServerHwnd, GWL_WNDPROC, (LONG)WindowProc);
  SetReg(IPC_NAME, (int)ServerHwnd);

  return 0;
};

extc int _export cdecl ODBG_Pluginmenu(int origin,char data[4096],void *item) {

	if(origin==PM_MAIN){
		strcpy(data,"0 &OllyServer");
		return 1;
	}
    
	return 0;

};

extc void _export cdecl ODBG_Pluginaction(int origin,int action,void *item) {
  if (origin==PM_MAIN) {
    switch (action) {
      case 0: MessageBox(0,"OllyServer - Running","",0); break;
      default: break;
    };
  };
};

extc void _export cdecl ODBG_Pluginreset(void) {
};

extc int _export cdecl ODBG_Pluginclose(void) {
  return 0;
};

extc void _export cdecl ODBG_Plugindestroy(void) {
	SetWindowLong(ServerHwnd, GWL_WNDPROC, (LONG)oldProc);
	DestroyWindow(ServerHwnd);
	SetReg(IPC_NAME,0);
};



//extc void _export cdecl ODBG_Pluginmainloop(DEBUG_EVENT *debugevent){}

