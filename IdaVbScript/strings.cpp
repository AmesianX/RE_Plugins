/*

 Sample how to integrate VB UI for IDA plugin
 http://sandsprite.com/CodeStuff/VB_Plugin_for_Olly.html

*/

#include <windows.h>  //define this before other headers or get errors 
#include <ida.hpp>
#include <idp.hpp>
#include <expr.hpp>
#include <bytes.hpp>
#include <loader.hpp>
#include <kernwin.hpp>
#include <name.hpp>
#include <auto.hpp>
#include <frame.hpp>
#include <dbg.hpp>
#include <area.hpp>
#include <stdio.h>
#include <search.hpp>
#include <xref.hpp>

#undef strcpy
 
IDispatch        *IDisp;

int StartPlugin(void);
int InitPlugin(void);

//if no use windows.h you can declare API fx manually like
//extern "C" int GetProcAddress(int h, char* fxName);
//extern "C" int GetModuleHandle(char* modName);


//Initialize.called once. PLUGIN_OK = unload+recall, PLUGIN_KEEP = keep in mem
int idaapi init(void)
{
  if ( inf.filetype == f_ELF ) return PLUGIN_SKIP;
  InitPlugin();
  return PLUGIN_KEEP;
}

//      Terminate.
void idaapi term(void)
{
	try{
		IDisp->Release();
	}
	catch(...){};

	CoUninitialize();
}

void idaapi run(int arg)
{
 
  StartPlugin();

}

char comment[] = "Misc functionality for IDA";
char help[] ="Contains lots of misc features and tools for IDA.\n\nWritten in VB by dzzie http://sandsprite.com";
char wanted_name[] = "dzzies IDA Plugin";
char wanted_hotkey[] = "Alt-0";

//Plugin Descriptor Block
plugin_t PLUGIN =
{
  IDP_INTERFACE_VERSION,
  0,                    // plugin flags
  init,                 // initialize
  term,                 // terminate. this pointer may be NULL.
  run,                  // invoke plugin
  comment,              // long comment about the plugin (status line or hint)
  help,                 // multiline help about the plugin
  wanted_name,          // the preferred short name of the plugin
  wanted_hotkey         // the preferred hotkey to run the plugin
};

int InitPlugin(){

    //Create an instance of our VB COM object, and execute
	//one of its methods so that it will load up and show a UI
	//for us, then it uses our other exports to access olly plugin API
	//methods

	CLSID      clsid;
	HRESULT	   hr;
    LPOLESTR   p = OLESTR("IdaVbSample.CPlugin");

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

	 OLECHAR *sMethodName = OLESTR("InitPlugin");
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
	 dispparams.rgvarg = &vararg[0];
	 dispparams.cArgs = 0;  // num of args function takes
	 dispparams.cNamedArgs = 0;

	 // and invoke the method
	 hr=IDisp->Invoke( dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, &retVal, NULL, NULL);

	 return 0;
}





int StartPlugin(){

    //Create an instance of our VB COM object, and execute
	//one of its methods so that it will load up and show a UI
	//for us, then it uses our other exports to access olly plugin API
	//methods

	CLSID      clsid;
	HRESULT	   hr;
    LPOLESTR   p = OLESTR("IdaVbSample.CPlugin");

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
	 dispparams.rgvarg = &vararg[0];
	 dispparams.cArgs = 0;  // num of args function takes
	 dispparams.cNamedArgs = 0;

	 // and invoke the method
	 hr=IDisp->Invoke( dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, &retVal, NULL, NULL);

	 return 0;
}





//Export API for the VB app to call and access IDA API data
//_________________________________________________________________
void __stdcall Jump      (int addr)  { jumpto(addr);           }
void __stdcall Refresh   (void)      { refresh_idaview();      }
int  __stdcall ScreenEA  (void)      { return get_screen_ea(); }
int  __stdcall NumFuncs  (void)      { return get_func_qty();  }
void __stdcall RemvName  (int addr)  { del_global_name(addr);  }
void __stdcall Setname(int addr, const char* name){ set_name(addr, name); }
void __stdcall AddComment(char *cmt){ generate_big_comment(cmt, COLOR_REGCMT);}
void __stdcall AddProgramComment(char *cmt){ add_pgm_cmt(cmt); }
void __stdcall AddCodeXRef(int start, int end){ add_cref(start, end, cref_t(fl_CN | XREF_USER) );}
void __stdcall DelCodeXRef(int start, int end){ del_cref(start, end, 1 );}
void __stdcall AddDataXRef(int start, int end){ add_dref(start, end, dref_t(dr_O | XREF_USER) );}
void __stdcall DelDataXRef(int start, int end){ del_dref(start, end );}
void __stdcall MessageUI(char *m){ msg(m);}
void __stdcall PatchByte(int addr, char val){ patch_byte(addr, val); }
void __stdcall PatchWord(int addr, int val){  patch_word(addr, val); }
void __stdcall DelFunc(int addr){ del_func(addr); }
int  __stdcall FuncIndex(int addr){ return get_func_num(addr); }
void __stdcall SelBounds( ulong* selStart, ulong* selEnd){ read_selection(selStart, selEnd);}
void __stdcall FuncName(int addr, char *buf, size_t bufsize){ get_func_name(addr, buf, bufsize);}
int  __stdcall GetBytes(int offset, void *buf, int length){ return get_many_bytes(offset, buf, length);}
void __stdcall Undefine(int offset){ autoMark(offset, AU_UNK); }
char __stdcall OriginalByte(int offset){ return get_original_byte(offset); }

int __stdcall ImageBase(void){

  netnode penode("$ PE header");
  ea_t loaded_base = penode.altval(-2);
  return loaded_base;

}

int __stdcall SetComment(unsigned int offset, char* comm){
	
	try{
		set_cmt(offset,comm,true);	
	}catch(...){
		return 0;
	}

	return 1;
}

int __stdcall GetComment(int offset, char* buf){ 

//	idaman ssize_t ida_export get_cmt(ea_t ea, bool rptble, char *buf, size_t bufsize);
	char b[255];

	unsigned int sz = get_cmt(offset,false,buf,512);
	
	//qsnprintf(b,254,"sz=%x",sz);
	//MessageBox(0,b,"",0);
	
	if(sz == -1){
		sz = get_cmt(offset,true,buf,512);
		//qsnprintf(b,254,"sz=%x",sz);
		//MessageBox(0,b,"",0);
	}

	return sz;

}

int __stdcall GetRComment(int offset, char* buf){ 
	int sz = get_cmt(offset,true,buf,512);
	return sz;
}


int __stdcall ProcessState(void){ return get_process_state(); }

int __stdcall FilePath(char *buf){ 
	int retlen=0;
	char *str;

	return get_input_file_path(buf,255);

}

int __stdcall RootFileName(char *buf){ 
	int retlen=0;
	char *str;

	return get_root_filename(buf,255);
	
}

void __stdcall HideEA(int offset){	set_visible_item(offset, false); }
void __stdcall ShowEA(int offset){	set_visible_item(offset, true); }


int __stdcall NextAddr(int offset){
   return nextaddr(offset);
}

int __stdcall PrevAddr(int offset){
    return prevaddr(offset); 
}



//not working?
void __stdcall AnalyzeArea(int startat, int endat){ /*analyse_area(startat, endat);*/}


//now working w/ labels
void __stdcall GetName(int offset, char* buf, int bufsize){

	get_true_name( BADADDR, offset, buf, bufsize );

	if(strlen(buf) == 0){
		func_t* f = get_func(offset);
		for(int i=0; i < f->llabelqty; i++){
			if( f->llabels[i].ea == offset ){
				int sz = strlen(f->llabels[i].name);
				if(sz < bufsize) strcpy(buf,f->llabels[i].name);
				return;
			}
		}
	}

}

//not workign to make code and analyze
void __stdcall MakeCode(int offset){
	 autoMark(offset, AU_CODE);
	 //analyse_area(offset, (offset+1) );
}


int __stdcall FunctionStart(int n){
	func_t *clsFx = getn_func(n);
	return clsFx->startEA;
}

int __stdcall FunctionEnd(int n){
	func_t *clsFx = getn_func(n);
	return clsFx->endEA;
}

int __stdcall FuncArgSize(int index){
		func_t *clsFx = getn_func(index);
		return clsFx->argsize ;
}

int __stdcall FuncColor(int index){
		func_t *clsFx = getn_func(index);
		return clsFx->color  ;
}

int __stdcall GetAsm(int addr, char* buf, int bufLen){

    flags_t flags;                                                       
    int sLen=0;

    flags = getFlags(addr);                        
    if(isCode(flags)) {                            
        generate_disasm_line(addr, buf, bufLen, GENDSM_MULTI_LINE );
        sLen = tag_remove(buf, buf, bufLen);  
    }

	return sLen;

}


//idaman ea_t ida_export find_text(ea_t startEA, int y, int x, const char *ustr, int sflag);
//#define SEARCH_UP       0x000		// only one of SEARCH_UP or SEARCH_DOWN can be specified
//#define SEARCH_DOWN     0x001
//#define SEARCH_NEXT     0x002

int __stdcall SearchText(int addr, char* buf, int search_type,int debug){

	char msg[500]={0};
	int y=0,x=0;
	int ret = find_text(addr,y,x,buf, search_type);
	
	if(debug==1){
		qsnprintf(msg,499,"ret=%x addr=%x search_type=%x",ret,addr,search_type);
		MessageBox(0,msg,"",0);
	}

	return ret;

}

int __stdcall GetRefsTo(int offset, int callback){

	//if(!callback) return 0;
	int count=0;
	int retVal=0;
	int ( __stdcall *lpVBFunction )(int,int); 
	
	_asm{
		mov eax, callback
		mov lpVBFunction, eax
	}

	xrefblk_t xb;
    for ( bool ok=xb.first_to(offset, XREF_ALL); ok; ok=xb.next_to() ){
		//if(xb.type == 1){
			retVal = (*lpVBFunction)(offset,xb.from  );
			count++;
			//if(retVal== -1) break;
		//}
    }
	
	return count;

}

int __stdcall GetRefsFrom(int offset, int callback){

	//this also returns jmp type xrefs not just call
	//there is always one back reference from next instruction 

	//if(!callback) return 0;
	int count=0;
	int retVal=0;
	int ( __stdcall *lpVBFunction )(int,int); 

	_asm{
		mov eax, callback
		mov lpVBFunction, eax
	}

	xrefblk_t xb;
    for ( bool ok=xb.first_from(offset, XREF_ALL); ok; ok=xb.next_from() ){
		//if(xb.iscode == 1){
			retVal = (*lpVBFunction)(offset,xb.to  );
			count++;
			//if(retVal== -1) break;
		//}
    }
	
	return count;

}




