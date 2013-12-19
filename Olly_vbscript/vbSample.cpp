// VERY IMPORTANT NOTICE: COMPILE THIS DLL WITH BYTE ALIGNMENT OF STRUCTURES
// AND UNSIGNED CHAR!

#include <windows.h>
#include <stdio.h>
#include <string.h>
#include <ole2.h>
#include "plugin.h"

HINSTANCE        hinst;
HWND             hwmain;
IDispatch        *IDisp;
ulong			 threadID;
int				 BpxHooking=0;
int				 RemoteSyncActive=0;
int				 DEBUGMODE = 1;

int  CreatePlugin(void);
void OnBreakPoint(int);
void FireSync(int,int);
int  RunPluginCmd(int);
 
void dbg(int x){
		char buf[12];
		memset(buf,0,12);
		sprintf(buf, "x=0x%x",x);
		MessageBox(0,buf,"",0);
}

BOOL WINAPI DllEntryPoint(HINSTANCE hi,DWORD reason,LPVOID reserved) {
  if (reason==DLL_PROCESS_ATTACH)
    hinst=hi;
  return 1;
};

extc int _export cdecl ODBG_Plugindata(char shortname[32]) {
  strcpy(shortname,"Olly VbScript//IDASync");
  return PLUGIN_VERSION;
};

extc int _export cdecl ODBG_Plugininit( int ollydbgversion,HWND hw,ulong *features) {
  if (ollydbgversion<PLUGIN_VERSION) return -1;
  
  hwmain=hw;
  Addtolist(0,0,"Olly VbScript//IDASync");
  CreatePlugin();

  return 0;
};

extc int _export cdecl ODBG_Pluginmenu(int origin,char data[4096],void *item) {
  switch (origin) {
    case PM_MAIN:
      strcpy(data,"0 &Script UI,|1 IDA Sync");
      return 1;
    default: break;
  };
  return 0;
};

extc void _export cdecl ODBG_Pluginaction(int origin,int action,void *item) {
  if (origin==PM_MAIN) { 
	 RunPluginCmd(action);
  };
};

extc void _export cdecl ODBG_Pluginreset(void) {
};

extc int _export cdecl ODBG_Pluginclose(void) {
  return 0;
};

extc void _export cdecl ODBG_Plugindestroy(void) {
	
	try{
		IDisp->Release();
	}
	catch(...){};

	CoUninitialize();

};

int ImageBaseForEA(int ea){

	t_table *mod_table;
	t_module *module;
	char buf[250];

	mod_table = (t_table *)Plugingetvalue(VAL_MODULES);

	for (int i=0; i<mod_table->data.n; i++){

        if ((module = (t_module *) Getsortedbyselection(&(mod_table->data), i)) == NULL){
            if(DEBUGMODE) Addtolist(0, 1, "OllySync> Unable to resolve module at index %d.", i);
            continue;
        }
		
		if(module->name == NULL) break;
		if( ea > module->base && ea < (module->base + module->size) ){
			return module->base;
		}

    }

	return 0;

}

int __stdcall ImageBaseAndNameForEA(int ea, int *outValBase, char* buf, int bufSize){

	t_table *mod_table;
	t_module *module;

	mod_table = (t_table *)Plugingetvalue(VAL_MODULES);

	for (int i=0; i<mod_table->data.n; i++){

        if ((module = (t_module *) Getsortedbyselection(&(mod_table->data), i)) == NULL){
            if(DEBUGMODE) Addtolist(0, 1, "OllySync> Unable to resolve module at index %d.", i);
            continue;
        }
		
		if(module->name == NULL) break;
		if( ea > module->base && ea < (module->base + module->size) ){
			*outValBase = module->base;
			int slen = strlen(module->name);
			if(slen < bufSize) strcpy(buf,module->name); else strcpy(buf,"Buf to Small");
			return 1;
		}

    }

	return 0;

}

extc void _export cdecl ODBG_Pluginmainloop(DEBUG_EVENT *debugevent) 
{
	t_status status;
	status = Getstatus();
	int ea=0;

	//TODO: query IDA instance found to make sure the rva we are FireSync on
	//      is in the same IDB as the module-> name other wise this is crazy
	//      before this was VA based which was ok cause jump would fail if invalid address...
	//      so now we support rebased dlls..we must specify IDB/module to make sure same...
	//      bleh so much work!
	//      real slick would be to enum open ida instances and select which one to send to
	//      based on the module name. so you could sync multiple ida's correctly if you were
	//      working on a dll system. you would need to switch to copydata though over sockets.

	if( debugevent && debugevent->dwDebugEventCode == EXCEPTION_DEBUG_EVENT)
	{
		EXCEPTION_DEBUG_INFO edi = debugevent->u.Exception;
		int ea = (int)edi.ExceptionRecord.ExceptionAddress   ;
		//int rva = ea - ImageBaseForEA(ea);

		//dbg(ea);

		if(edi.ExceptionRecord.ExceptionCode == EXCEPTION_BREAKPOINT){
			
			
			if(BpxHooking == 1)        OnBreakPoint(ea);
			if(RemoteSyncActive == 1 ) FireSync(0,ea);

		}
		else if(edi.ExceptionRecord.ExceptionCode == EXCEPTION_SINGLE_STEP ){

			if(BpxHooking == 1)        OnBreakPoint(ea);
			if(RemoteSyncActive == 1 ) FireSync(0,ea);
		
		}
	}
}


int CreatePlugin(){

    //Create an instance of our VB COM object, and execute
	//one of its methods so that it will load up and show a UI
	//for us, then it uses our other exports to access olly plugin API
	//methods

	CLSID      clsid;
	HRESULT	   hr;
    LPOLESTR   p = OLESTR("OllyVBScript.CPlugin");

    hr = CoInitialize(NULL);

	 hr = CLSIDFromProgID( p , &clsid);
	 if( hr != S_OK  ){
		 MessageBox(0,"Failed to get Clsid from string\n","",0);
		 return 0;
	 }

	 // create an instance and get IDispatch pointer
	 hr =  CoCreateInstance( clsid,
							 NULL,
							 CLSCTX_INPROC_SERVER,
							 IID_IDispatch  ,
							 (void**) &IDisp
						   );

	 if ( hr != S_OK )
	 {
	   MessageBox(0,"CoCreate failed","",0);
	   return 0;
	 }

	 return 1;
}


int RunPluginCmd(int arg){
	
	 HRESULT	   hr;

	 OLECHAR *sMethodName = OLESTR("DoPluginAction");
	 DISPID  dispid; // long integer containing the dispatch ID

	 // Get the Dispatch ID for the method name
	 hr=IDisp->GetIDsOfNames(IID_NULL,&sMethodName,1,LOCALE_USER_DEFAULT,&dispid);
	 if( FAILED(hr) ){
	    MessageBox(0,"GetIDS failed","",0);
		return 0;
	 }

	 DISPPARAMS dispparams;
	 VARIANTARG vararg[1]; //function takes one argument
	 VARIANT    retVal;

	 VariantInit(&vararg[0]);

	 vararg[0].vt = VT_I4 ;
	 vararg[0].intVal = arg;

	 dispparams.rgvarg = &vararg[0];
	 dispparams.cArgs = 1;  // num of args function takes
	 dispparams.cNamedArgs = 0;

	 // and invoke the method
	 hr=IDisp->Invoke( dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, &retVal, NULL, NULL);

	 return 0;
}

void OnBreakPoint(int addr){
	 
	 OLECHAR *sMethodName = OLESTR("BreakPoint");
	 DISPID  dispid; // long integer containing the dispatch ID
	 HRESULT hr;

	 // Get the Dispatch ID for the method name
	 hr=IDisp->GetIDsOfNames(IID_NULL,&sMethodName,1,LOCALE_USER_DEFAULT,&dispid);
	 if( FAILED(hr) ){
	    MessageBox(0,"GetIDS failed","",0);
		return  ;
	 }

	 DISPPARAMS dispparams;
	 VARIANTARG vararg[1]; //function takes one argument
	 VARIANT    retVal;
	
	 VariantInit(&vararg[0]);
	 
	 vararg[0].vt = VT_I4 ;
	 vararg[0].intVal = addr;

	 dispparams.rgvarg = &vararg[0];
	 dispparams.cArgs = 1;  // num of args function takes
	 dispparams.cNamedArgs = 0;

	//dbg(1);
	 // and invoke the method
	 hr=IDisp->Invoke( dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, &retVal, NULL, NULL);
	//dbg(2);
}

void FireSync(int event, int arg){
	 
	 OLECHAR *sMethodName = OLESTR("FireSync");
	 DISPID  dispid; // long integer containing the dispatch ID
	 HRESULT hr;

	 if(DEBUGMODE) Addtolist(0, 1, "OllySync.FireSync> Event %d Arg: 0x%x", event, arg);

	 // Get the Dispatch ID for the method name
	 hr=IDisp->GetIDsOfNames(IID_NULL,&sMethodName,1,LOCALE_USER_DEFAULT,&dispid);
	 if( FAILED(hr) ){
	    MessageBox(0,"GetIDS failed","",0);
		return  ;
	 }

	 DISPPARAMS dispparams;
	 VARIANTARG vararg[2]; //function takes one argument
	 VARIANT    retVal;

	 VariantInit(&vararg[0]);

	 vararg[0].vt = VT_I4 ;
	 vararg[0].intVal = arg;    //second arg to function

	 vararg[1].vt = VT_I4 ;
	 vararg[1].intVal = event;  //first arg to function

	 dispparams.rgvarg = &vararg[0];
	 dispparams.cArgs = 2;  // num of args function takes
	 dispparams.cNamedArgs = 0;
	
	 //dbg(3);
	 // and invoke the method
	 hr=IDisp->Invoke( dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, &retVal, NULL, NULL);
	 //dbg(4);
}


//Export API for the VB app to call and access Olly Plugin API data
//_________________________________________________________________
int __stdcall GetModuleBaseAddress(void){
  struct t_dump *currentmodule;
  
  currentmodule = (t_dump *)Plugingetvalue(VAL_CPUDASM);
  return (int)currentmodule->base;
}

int __stdcall GetModuleSize(void){
  struct t_dump *currentmodule;
  
  currentmodule = (t_dump *)Plugingetvalue(VAL_CPUDASM);
  return (int)currentmodule->size;
}

int __stdcall GetName(int *offset, int *type, LPSTR pszString){

  char name[256];
  int retVal=0;
  
  retVal = Findname(*offset, *type, name) ;

  if(retVal > 0) strcpy(pszString, name);
	
  return retVal;

}

//int Insertname(ulong addr,int type,char *name);
int __stdcall SetName(int *offset, int *type, LPSTR pszString){
  return Insertname(*offset, *type, pszString) ;
}


//ulong Findprocbegin(ulong addr);
int __stdcall ProcBegin(int *addr){
	return Findprocbegin(*addr);
}

//ulong Findprocend(ulong addr);
int __stdcall ProcEnd(int *addr){
	return Findprocend(*addr);
}

//ulong Findprevproc(ulong addr);
int __stdcall PrevProc(int *addr){
	return Findprevproc(*addr);
}

//ulong Findnextproc(ulong addr);
int __stdcall NextProc(int *addr){
	return Findnextproc(*addr);
}

//void Animate(int animation);
void __stdcall DoAnimate(int *animation){
	Animate(*animation);
}

//int Deletehardwarebreakbyaddr(ulong addr);
int __stdcall DelHdwBP(int *addr){
	return Deletehardwarebreakbyaddr(*addr);
}

//int Setbreakpoint(ulong addr,ulong type,uchar cmd);
int __stdcall SetBpx(int *addr, int *type, char *cmd){
	return Setbreakpoint(*addr, *type, *cmd);
}

//int Sethardwarebreakpoint(ulong addr,int size,int type);
int __stdcall SetHdwBpx(int *addr, int *size, int *type){
	return Sethardwarebreakpoint(*addr, *size, *type);
}



//ulong Readmemory(void *buf,ulong addr,ulong size,int mode);
int __stdcall ReadMem(char *buf, int *addr, int *size, int *mode){
	return Readmemory(buf, *addr, *size, *mode);
}


//ulong Writememory(void *buf,ulong addr,ulong size,int mode);
int __stdcall WriteMem(char *buf, int*addr, int *size, int *mode){
	return Writememory(buf, *addr, *size, *mode);
}

//void Redrawdisassembler(void);
void __stdcall RefereshDisasm(void){
	Redrawdisassembler();
}

//ulong Getcputhreadid(void);
int __stdcall CurThreadID(void){
	return Getcputhreadid();
}

//ulong Followcall(ulong addr);
int __stdcall Follow(int *addr){
	return Followcall(*addr);
}

//int Go(ulong threadid,ulong tilladdr,int stepmode,int givechance,int backupregs);
int __stdcall DoGo(int *threadid, int *tilladdr, int *stepmode, int *givechance, int *backupargs){
	return Go(*threadid, *tilladdr, *stepmode, *givechance, *backupargs);
}

int __stdcall GetStat(void){
	return Getstatus();
}

//ulong Disasm(char *src,        ulong srcsize,   ulong srcip,     char *srcdec,
//             t_disasm *disasm, int disasmmode,  ulong threadid);
int __stdcall GetAsm(char *src, int *srcsize, int *srcip, LPSTR pszString){
	t_disasm td;
	int codeLen;
	codeLen = Disasm((unsigned char*)src, *srcsize, *srcip, 0,  &td,DISASM_CODE, Getcputhreadid());
	strcpy(pszString, td.result);
	return codeLen;
	
}


//ulong Readcommand(ulong ip,char *cmd);
int __stdcall GetByteCode(int *ip, LPSTR pszString){
	Readcommand(0, pszString); //invalidate cache
	return Readcommand(*ip, pszString);
}


//int Assemble(char *cmd,ulong ip,t_asmmodel *model,int attempt,int constsize,char *errtext);
int __stdcall Asm(LPSTR cmd, int *addr, LPSTR errBuffer, LPSTR codeBuffer){
    t_asmmodel t;
	Assemble( cmd, *addr, &t, 0, 0, errBuffer);
	strcpy(codeBuffer, (const char*)t.code);
	return t.length ;
	
}

void __stdcall EnableBpxHook(){
	BpxHooking = 1;
}

void __stdcall DisableBpxHook(){
	BpxHooking = 0 ;
}

int __stdcall GetEIP(void){
	t_thread *th = Findthread(Getcputhreadid());
	t_reg tr;
	tr =  th->reg; 
	return tr.ip ;
}

void __stdcall SetEIP(int v){
	t_thread *th = Findthread(Getcputhreadid());
	th->reg.ip = v;
}

int __stdcall GetRegister(int i){
	t_thread *th = Findthread(Getcputhreadid());
	t_reg tr;
	tr =  th->reg; 
    return (int)tr.r[i];
}

void __stdcall SetRegister(int rIndex, int v ){
	t_thread *th = Findthread(Getcputhreadid());
	th->reg.r[rIndex] = v ;
}

void __stdcall RedrawCpu(void){
	Setcpu(0,0,0,0, CPU_REDRAW);
}

int __stdcall ConfigValue(int request){
	return  Plugingetvalue(request);
}

//int OpenEXEfile(char *path,int dropped);
int __stdcall OpenEXE(char *path){
	return OpenEXEfile(path,0);
}

//int Decodeaddress(ulong addr,ulong base,int addrmode,char *symb,int nsymb,char *comment);
int __stdcall DecodeAddr(int *addr, int *base, int*addrmode, char*buf, int *buflen, char *combuf){
		return Decodeaddress(*addr, *base, *addrmode, buf, *buflen, combuf);
}

//int Decodeascii(ulong addr,char *s,int len,int mode);
int __stdcall DecodeAsc(int *addr, char *buf, int *len){
	return Decodeascii(*addr, buf, *len, DASC_ASCII);	
}

//int Decodeunicode(ulong addr,char *s,int len);
int __stdcall DecodeUni(int *addr, char *buf, int *len){
	return Decodeunicode(*addr, buf, *len );	
}

//ulong Findimportbyname(char *name,ulong addr0,ulong addr1);
int __stdcall FindImportName(char *buf, int *startAddr,int *endAddr){
		return Findimportbyname(buf, *startAddr, *endAddr);
}

//int Deletehardwarebreakbyaddr(ulong addr);
void __stdcall DelHwrBPXAddr(int*addr){
	Deletehardwarebreakbyaddr(*addr);
}

//int Deletehardwarebreakpoint(int index);
void __stdcall DelHwrBPXIndex(int *index){
	Deletehardwarebreakpoint(*index);
}

void __stdcall SetSyncFlag(int active){ RemoteSyncActive = active; } 


