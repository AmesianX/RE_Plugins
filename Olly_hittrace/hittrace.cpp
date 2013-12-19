/*

  Author: David Zimmer <david@idefense, dzzie@yahoo.com>

  Notes: this is ugly and was done as quick as possible but it works
         if you need more than 50 bpx recompile;

		 THIS APP USES DLL NAME LISTED IN MODULES EXPORT TABLE FOR
		 STRING MATCHING. IF DLL NAME IS DIFFERENT IT WILL LOOK LIKE
		 THIS APP ISNT WORKING.

 
  VERY IMPORTANT NOTICE: COMPILE THIS DLL WITH BYTE ALIGNMENT OF STRUCTURES
  AND UNSIGNED CHAR!

  Bugfixes:
	9/15/06 -
			fixed potential bof
			fixed errs in loading saved lists & added msgs if save err

    10/22/2010
	      - added log to file feature
		  - cleaned up code
		  - fixed a bad bug where stale data could be logged !
		  - added a comment feature (in ui and extract from disasm (ui takes precedence))
		  - added the A prefix to take the expression and do an ascii dump
		  - added Hxx prefix to do a hexdump where xx is decimal length to dump
			

*/

#include <windows.h>
#include <stdio.h>

#include <string.h>
#include <StdDef.H> // offsetof()
#include "resource.h"
#include "plugin.h"

struct BpxEntry{
	char			  dllName[50];
	char			  LogExp[50];
	char              comment[150];  
	unsigned int      rvaBpx;
	unsigned int      hitcnt;
	unsigned int      va;
	bool              enabled; //should check olly to make sure not deleted?
};

BpxEntry Entries[50];  

int  numEntries           =  -1;
int  m_hProcess		      =  0;
bool debug			      =  false;
int  my_timer		      =  0;
int  abort_trace_index    =  -1;
char szFileName[MAX_PATH] = {0};
int  active_bp_va         = 0;
int  msgbox_debug         = 1;
FILE *fp                  = NULL;

char* logfile = "c:\\hittrace.log";


HINSTANCE        hinst;
HWND             hwmain;
HINSTANCE        myHnd;

int MyExpression(char*,int);
int GetRegister(int,int);

void ShowFileLog(){

	if(fp){
		fclose(fp);
		fp=NULL;
	}

	WinExec("notepad.exe c:\\hittrace.log", 1);

	for(int i=0;i<5;i++){
		Sleep(200);
		HWND h = FindWindow("Notepad", "hittrace.log - Notepad");
		if(h!=0){ 
			SetForegroundWindow(h); 
			DeleteFile("c:\\hittrace.log");
			break;
		}
	}

}

char* HexDump(char* bufIn, int len){
	int mlen = (len+2)*3;

	char tmp[10];
	char* bufOut = (char*)malloc(mlen);
	memset(bufOut,0, mlen);

	for(int i=0;i<len;i++){
		memset(tmp,0,9);
		sprintf((char*)&tmp,"%02X ", bufIn[i]);
		strcat(bufOut,tmp);
	}

	return bufOut;
}


int IntStrLen(int i){
	char x[50] = {0};
	memset(x,0,49);
	sprintf(x,"%d",i);
	return strlen(x);
}

void logmsg(int red, const char *format, ...)
{
	DWORD dwErr = GetLastError();
		
	if(format){
		char buf[1024];
		va_list args; 
		va_start(args,format); 
		try{
			_vsnprintf(buf,1024,format,args);
			Addtolist(0,red,buf);

			if(fp==NULL) fp = fopen(logfile, "w");
			if(fp){
				fwrite(buf, 1, strlen(buf), fp);
				fwrite("\r\n",1,2,fp);
			}

			if(debug){
				if(msgbox_debug > 0){
					MessageBox(hwmain,buf,"",0);
				}else{
					//todo log to file
				}
			}

		}
		catch(...){}
	}

	SetLastError(dwErr);
}


void strlower(char* x){
	if( (int)x ){
		int y = strlen(x);
		for(int i=0;i<y;i++){
			x[i]=tolower(x[i]);
		}
	}
}

void FillList(HWND hList){
	char buf[200];
	SendMessage(hList,LB_RESETCONTENT,0,0);
	for(int i=0;i<=numEntries;i++){
		sprintf(buf, "%d:          %s          %x          %s          %s", i, Entries[i].dllName, Entries[i].rvaBpx , Entries[i].LogExp, Entries[i].comment );
		SendMessage(hList,LB_ADDSTRING,0,(int)&buf);
	}
}

bool AddEntry(char* mod, int rva, char* logexp, char* cmt){

	if(numEntries > 49) return false;
	if(strlen(mod) > 49) mod[49]=0;
	if(strlen(cmt) > 149) mod[149]=0;

	numEntries++;
	strncpy(Entries[numEntries].dllName ,mod,strlen(mod));
	strncpy(Entries[numEntries].LogExp  ,logexp,49);
	strncpy(Entries[numEntries].comment ,cmt,149);

	Entries[numEntries].rvaBpx = rva;
	Entries[numEntries].va =0;
	Entries[numEntries].enabled =false;
	Entries[numEntries].hitcnt =0;
	
	return true;

}


void RemoveEntry(int index){
	
	bool compact = false;

	for(int i=0;i<=numEntries;i++){
top:
		if(compact){
			if(i==numEntries) memset((void*)&Entries[i],0,sizeof(BpxEntry));
			else Entries[i] = Entries[i+1];
		}
		else if(i==index){
				compact = true;
				goto top;
		}
	}

	if(compact) numEntries--;

}

void SetTop(HWND h, bool top=true){
	HWND flag = top ? HWND_TOPMOST : HWND_NOTOPMOST ;	
	SetWindowPos(h,flag,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE);
}
		
BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{

    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
    }
	
	//DLL instance MUST be used to load embedded resources.
	myHnd = (HINSTANCE)hModule;

    return TRUE;
}

//basically ripped from pedrams olly bp man ;)
int FindModuleFromName(char* modName){

	t_table *mod_table;
	t_module *module;
	char buf[250];

	mod_table = (t_table *)Plugingetvalue(VAL_MODULES);

	for (int i=0; i<mod_table->data.n; i++){

        if ((module = (t_module *) Getsortedbyselection(&(mod_table->data), i)) == NULL){
            if(debug) Addtolist(0, 1, "HitTrace> Unable to resolve module at index %d.", i);
            continue;
        }
		
		if(module->name == NULL) break;
		strncpy(buf, module->name, 249);
		strlower(buf);

        if (strstr(module->name, modName) != 0){
		     return module->base;
        }
    }

	return 0;

}

// This is the message loop
BOOL CALLBACK MyDlgProc(HWND hWnd,UINT uMsg,WPARAM wParam,LPARAM lParam)
{

	char buf [200]={0},modName[50]={0},sRva[50] = {0},sLog[50] = {0},sCmt[150] = {0} ;
    unsigned int i=0, index=0, rva=0, hexValue=0,base=0;
	int ret=0;
	FILE* dat;

	HWND hName = GetDlgItem(hWnd, IDC_EDIT1);
	HWND hRva  = GetDlgItem(hWnd, IDC_EDIT2);
	HWND hAbort= GetDlgItem(hWnd, IDC_EDIT3);
	HWND hLog  = GetDlgItem(hWnd, IDC_EDIT4);
	HWND hCmt  = GetDlgItem(hWnd, IDC_EDIT5);
	HWND hList = GetDlgItem(hWnd, IDC_LIST2);
	HWND hCheck= GetDlgItem(hWnd, IDC_CHECK1);


	switch(uMsg)
	{
		case WM_INITDIALOG :
			
			FillList(hList);
			if(debug) SendMessage(hCheck, BM_SETCHECK, BST_CHECKED ,0);
			if(abort_trace_index >=0){
				sprintf(sLog,"%d",abort_trace_index);
				SetWindowText(hAbort,sLog);
			}
			SetTop(hWnd);
			break;

		case WM_COMMAND:
			switch(LOWORD(wParam))
			{
				case IDC_BUTTON1: //delete item from listbox

					index = SendMessage(hList, LB_GETCURSEL, 0,0 );
					if(index== LB_ERR) break;

					SendMessage(hList,LB_GETTEXT,index, (int)&buf);
					index = atoi(buf); //now EntryIndex
					RemoveEntry(index);
					FillList(hList);
					
					break;
			
				case IDC_BUTTON2: //OK button ID
					
	
					GetWindowText(hName, modName,49);
					GetWindowText(hRva, sRva,49);
					GetWindowText(hLog, sLog,49);
					GetWindowText(hCmt, sCmt,149);

					strlower(modName);
					rva = strtol(sRva, NULL, 16);
					
					if(rva==0 || strlen(modName) == 0){
						MessageBox(0,"Invalid Input","",0);
						break;
					}

					if( AddEntry(modName, rva, sLog, sCmt) ){
						sprintf(buf, "%d:          %s          %x          %s          %s", numEntries, modName, rva, sLog, sCmt);
						SendMessage(hList,LB_ADDSTRING, 0, (int)&buf);
					}
					else{
						MessageBox(0,"Failed to add (maxed out)","",0);
					}
						

					break;

				case IDC_BUTTON3: //save list

					try{
						SetTop(hWnd,false);
						ret = Browsefilename("Save File As",(char*)&szFileName,".htl",0x80);
						SetTop(hWnd);

						if(debug) Addtolist(0,0,"Saving to file: %s", szFileName);
						if(ret != TRUE) break;
						dat = fopen(szFileName, "wb");
						fwrite((void*)Entries, sizeof(Entries),1,dat);
						fwrite((void*)&numEntries, 4,1,dat);
						fclose(dat);
					}
					catch(...){
						MessageBox(0,"Error Saving List","Error",0);
					}

					break;

				case IDC_BUTTON4: //load saved

					try{
						
						SetTop(hWnd,false);
						ret = Browsefilename("Open File",(char*)&szFileName,".htl;",0);
						SetTop(hWnd);

						if(debug) Addtolist(0,0,"Saving to file: %s", szFileName);
						if(ret != TRUE) break;
						if(debug) Addtolist(0,0,"Loading file: %s", szFileName);
						dat = fopen(szFileName, "rb");
						fread((void*)Entries, sizeof(Entries),1,dat);
						fread((void*)&numEntries, 4,1,dat);
						if(numEntries > 49) numEntries=49;

						for(i=0;i<numEntries;i++){
							Entries[i].enabled = false ;
						}

						fclose(dat);
						FillList(hList);
					}
					catch(...){
						MessageBox(0,"Error Loading List","Error",0);
					};

					break;

				case IDC_BUTTON5: //clear list

					memset( (void*)Entries, 0, sizeof(Entries));
					numEntries = -1;
					FillList(hList);
					break;


				case IDC_BUTTON6: //begin hit trace

					if(numEntries == -1){
						MessageBox(0,"You have not set any bpx","",0);
						break;
					}

					GetWindowText(hAbort, sRva,49);
					if(strlen(sRva)==0){
						abort_trace_index = -1;
					}
					else{
						abort_trace_index = atoi(sRva);
					}

					for(ret=0;ret <=numEntries;ret++){
						//we have already calculated the va so set bpx
						if(!Entries[ret].enabled && Entries[ret].va > 0){
							if(debug) Addtolist(0,0,"Not enabled va set");
							Setbreakpoint(Entries[ret].va , TY_ACTIVE,0);
							Entries[ret].enabled =false;
						}
						//breakpoint is in main module 
						else if(!Entries[ret].enabled && strcmp(Entries[ret].dllName, "main")==0){
							if(debug) Addtolist(0,0,"Bp in Main");
							base=Plugingetvalue(VAL_MAINBASE);
							Entries[ret].va = base + Entries[ret].rvaBpx;
							Entries[ret].enabled = true;
							Setbreakpoint( Entries[ret].va ,  TY_ACTIVE, 0);
						}
						//we have not calculated the va
						else if(!Entries[ret].enabled && Entries[ret].va == 0){
							if(debug) Addtolist(0,0,"Not enabled va 0");
							i = FindModuleFromName(Entries[ret].dllName);
							if(i>0){
								Entries[ret].va = i+Entries[ret].rvaBpx ;
								Entries[ret].enabled = true;
								Setbreakpoint( Entries[ret].va ,  TY_ACTIVE, 0);
							}
						}
						
					}

					Go(0,0,STEP_RUN,0,0);
					EndDialog(hWnd,0);
					break;

				case IDC_BUTTON7: //show results

					Suspendprocess(1);

					Addtolist(0,1,"Dumping hit trace results");
					for(ret=0;ret <=numEntries;ret++){
						if(Entries[ret].hitcnt > 0 ){
							Addtolist(0,0,"HitTrace %s.0x%x  = %d", Entries[ret].dllName, Entries[ret].rvaBpx , Entries[ret].hitcnt  );
						}
					}
					Addtolist(0,1,"Trace results done");
					Createlistwindow();
					ShowFileLog();
					break;


				case IDC_BUTTON8: //disable hit trace

					Suspendprocess(1);

					for(ret=0;ret <= numEntries;ret++){
						Setbreakpoint(Entries[ret].va , TY_DISABLED,0);
						Entries[ret].enabled =false;
					}


					break;

				case IDC_BUTTON9: //reset counts
					
					Suspendprocess(1);
					for(ret=0;ret < numEntries;ret++){
						Entries[ret].hitcnt = 0 ;
					}
					break;

				case IDC_BUTTON10: //edit selected

					index = SendMessage(hList, LB_GETCURSEL, 0,0 );
					if(index== LB_ERR) break;

					SendMessage(hList,LB_GETTEXT,index, (int)&buf);
					index = atoi(buf); //now EntryIndex
					
					char tmp[20];
					sprintf(tmp, "%x", Entries[index].rvaBpx);
					SetWindowText(hRva, tmp);
					SetWindowText(hCmt, Entries[index].comment);
					SetWindowText(hName, Entries[index].dllName);
					SetWindowText(hLog , Entries[index].LogExp);

					RemoveEntry(index);
					FillList(hList);

					break;

				case IDC_BUTTON11: //show counts
	
					char buf2[300];
					FILE* fp = fopen("c:\\hitcounts.txt","w");
					if(fp==NULL) break;
					
					strcpy(buf2, "HitCount     RVA       Comment\r\n");
					fwrite(buf2, 1, strlen(buf2), fp);

					for(i=0;i<=numEntries;i++){
						sprintf(buf2, "%04d   %8s.%04x  %s\r\n", Entries[i].hitcnt, Entries[i].dllName , Entries[i].rvaBpx , Entries[i].comment);
						fwrite(buf2,1,strlen(buf2), fp);
					}

					fflush(fp);
					fclose(fp);
					WinExec("notepad.exe c:\\hitcounts.txt", 1);

					for(i=0;i<5;i++){
						Sleep(200);
						HWND h = FindWindow("Notepad", "hitcounts.txt - Notepad");
						if(h!=0){
							SetForegroundWindow(h); break; 
							DeleteFile("c:\\hitcounts.txt");
						}
					}

					break;



			}

			debug=false;
			if(SendMessage(hCheck,BM_GETCHECK,0,0) == BST_CHECKED) debug = true;

			break;

		case WM_CLOSE:

			debug=false;
			if(SendMessage(hCheck,BM_GETCHECK,0,0) == BST_CHECKED) debug = true;
			EndDialog(hWnd,0);
			break;

	}

	return FALSE;

}



BOOL WINAPI DllEntryPoint(HINSTANCE hi,DWORD reason,LPVOID reserved) {
  if (reason==DLL_PROCESS_ATTACH) hinst=hi;
  return 1;
};

extc int _export cdecl ODBG_Plugindata(char shortname[32]) {
  strcpy(shortname,"HitTrace");
  return PLUGIN_VERSION;
};

extc int _export cdecl ODBG_Plugininit( int ollydbgversion,HWND hw,ulong *features) {
  if (ollydbgversion<PLUGIN_VERSION) return -1;
  Addtolist(0,0,"Initilizing HitTrace Plugin - David Zimmer (dzzie@yahoo.com)");
  hwmain=hw;
  return 0;
};

extc int _export cdecl ODBG_Pluginmenu(int origin,char data[4096],void *item) {

  switch (origin) {

		case PM_MAIN:
			  strcpy(data,"0 HitTrace Plugin, 1 About");
			  return 1;

		default: break;
  };

  return 0;

};

extc void _export cdecl ODBG_Pluginaction(int origin,int action,void *item) {
  if (origin==PM_MAIN) {
    switch (action) {
      case 0:
			 DialogBox( myHnd,MAKEINTRESOURCE(IDD_FORMVIEW),0,(DLGPROC)MyDlgProc); //modal
			 break;
	  case 1:
			 MessageBox(0,"HitTrace- David Zimmer <dzzie@yahoo.com>\n\n"
						   "Usage: Enter the dllname and the rva of the\n"
						   "code you want to hit trace\n\n"
						   "Optionally you can enter an expression to log\n"
						   "when the bpx is hit, click Begin to start\n\n"
						   "(You can also enter a index of bpx to stop at)","",0);
			 break;
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
};



// PURPOSE  : Retrieves the DLL module name for a given file handle of a
//            the module.  Reads the module name from the EXE header.
// from MS DEB Example in SDK I think
int GetModuleFileNameFromHeader(int hProcess, HANDLE hFile, 
								  int BaseOfDll, char* lpszPath,
								  int bufSize)
{
  #define IMAGE_SECOND_HEADER_OFFSET    (15 * sizeof(ULONG)) // relative to file beginning
  #define IMAGE_BASE_OFFSET             (13 * sizeof(DWORD)) // relative to PE header base
  #define IMAGE_EXPORT_TABLE_RVA_OFFSET (30 * sizeof(DWORD)) // relative to PE header base
  #define IMAGE_NAME_RVA_OFFSET         offsetof(IMAGE_EXPORT_DIRECTORY, Name)
	

  WORD   DosSignature;
  DWORD  NtSignature;
  DWORD  dwNumberOfBytesRead = 0;
  DWORD  PeHeader;
  int* ImageBase, ExportTableRVA, NameRVA;

  //-- verify that the handle is not 0L
  if( !hFile ) {
    strcpy(lpszPath,"Invalid File Handle");
    return( 0 );
  }

  //-- verify that the handle is for a disk file
  if( GetFileType(hFile) != FILE_TYPE_DISK ) {
    strcpy(lpszPath,"Invalid File Type");
    return( 0 );
  }

  //-- Extract the filename from the EXE header
  SetFilePointer(hFile, 0L, 0L, FILE_BEGIN );
  ReadFile( hFile, &DosSignature, sizeof(DosSignature), &dwNumberOfBytesRead,(LPOVERLAPPED) 0L);

  //-- verify DOS signature found
  if( DosSignature != IMAGE_DOS_SIGNATURE ) {
    sprintf(lpszPath,"Bad MZ Signature: 0x%x", DosSignature );
    return( 0 );
  }

  SetFilePointer( hFile, IMAGE_SECOND_HEADER_OFFSET, (LPLONG) 0L,FILE_BEGIN );
  ReadFile( hFile, &PeHeader, sizeof(PeHeader), &dwNumberOfBytesRead,(LPOVERLAPPED) 0L );
  SetFilePointer( hFile, PeHeader, (LPLONG) 0L, FILE_BEGIN );
  ReadFile( hFile, &NtSignature, sizeof(NtSignature), &dwNumberOfBytesRead,(LPOVERLAPPED) 0L);

  //-- verify Windows NT (PE) signature found
  if( NtSignature != IMAGE_NT_SIGNATURE ) {
    sprintf(lpszPath,"Bad PE Signature: 0x%x",DosSignature);
    return( 0 );
  }

  SetFilePointer( hFile, PeHeader + IMAGE_BASE_OFFSET, (LPLONG) 0L,FILE_BEGIN );
  ReadFile( hFile, &ImageBase, sizeof(ImageBase), &dwNumberOfBytesRead, (LPOVERLAPPED) 0L);
  SetFilePointer( hFile, PeHeader + IMAGE_EXPORT_TABLE_RVA_OFFSET,(LPLONG) 0L, FILE_BEGIN );
  ReadFile( hFile, &ExportTableRVA, sizeof(ExportTableRVA), &dwNumberOfBytesRead, (LPOVERLAPPED) 0L);

  //-- now read from the virtual address space in the process
  if (!ReadProcessMemory( (HANDLE)hProcess,
         (LPVOID) (BaseOfDll + ExportTableRVA + IMAGE_NAME_RVA_OFFSET),
         &NameRVA, sizeof(NameRVA), &dwNumberOfBytesRead ) ||
      dwNumberOfBytesRead != sizeof(NameRVA))
  {
      strcpy(lpszPath,"Access Denied!");
      dwNumberOfBytesRead = 0;
  }
  else if( !ReadProcessMemory( (HANDLE) hProcess,
              (LPVOID) (BaseOfDll + NameRVA), lpszPath, bufSize, &dwNumberOfBytesRead ))
  {
     strcpy(lpszPath,"Access Denied!");
  }

  return( dwNumberOfBytesRead );
}





VOID CALLBACK ShowFileLog_TimerProc(HWND hwnd,UINT uMsg,UINT_PTR idEvent,DWORD dwTime){

	KillTimer((HWND)0, my_timer);
	ShowFileLog(); 

}



	
void HandleTracerBp(int i){
	
	char buf[200];
	char comment[256];
	bool asciiDump=false;
	bool hexDump = false;
	char* exp = NULL;
	int cnt=0;
    int hasComment=0;
    int dumpLen=0;

	t_result result;
	memset(&result, 0, sizeof(t_result));
	memset(buf,0,200);
	memset(comment,0,256);

	if(debug) logmsg(1,"In OnTracer_BreakPoint i=%d", i);
		
	exp = Entries[i].LogExp;
	
	//the comment set in the ui takes precedence over the disasm comment
	if( strlen(Entries[i].comment) > 0 ){
		strcpy(comment, Entries[i].comment);
		hasComment = 1;
	}
	else{ //no comment set in ui, check to see if there is one in the disasm..
		hasComment = Findname(Entries[i].va, NM_COMMENT, comment) ;
	}

	//check to see if the ascii dump prefix is added to this expression..
	//if so strip it so we can pass real expression to Expression()
	if(Entries[i].LogExp[0]=='A' || Entries[i].LogExp[0]=='a'){
		cnt=1;
		while(Entries[i].LogExp[cnt]==' ') cnt++;
		asciiDump = true;
		exp = (char*)&Entries[i].LogExp[cnt];
	}

	if(debug) logmsg(1,"In OnTracer_BreakPoint before hexdump");

	if(Entries[i].LogExp[0]=='H' || Entries[i].LogExp[0]=='h'){ //hexdump
		cnt=1;
		try{
			dumpLen = atoi((char*)&Entries[i].LogExp[cnt]);
		}catch(...){
			if(debug) logmsg(1,"caught error in atoi");
			dumpLen=0;
		}

		if(dumpLen > 0){
			hexDump = true;
			cnt += IntStrLen(dumpLen);
			while(Entries[i].LogExp[cnt]==' ') cnt++;
			exp = (char*)&Entries[i].LogExp[cnt];
		}
	}
	
	if(debug)logmsg(1,"asciidump=%d hexDump=%d  exp ='%s' cnt=%d  dumpLen=%d ", (int)asciiDump, (int)hexDump, exp, cnt, dumpLen);

	int n = Expression(&result,exp,0,0,NULL,0,0,Getcputhreadid());
	
	if(debug) logmsg(1,"type = %x  u = %x, ascii=%x", result.dtype, result.u, (int)result.value);

	try{
		if (n>0 || result.type!=DEC_UNKNOWN) {
						
			if(asciiDump){ 
				Readmemory(&buf, result.u, 199, MM_RESTORE);
				if(hasComment>0){
					logmsg(1,"%s.0x%x  %s = %s  (%s)",Entries[i].dllName , Entries[i].rvaBpx , Entries[i].LogExp, buf,comment );
				}else{
					logmsg(1,"%s.0x%x  %s = %s ",Entries[i].dllName , Entries[i].rvaBpx , Entries[i].LogExp, buf);
				}
			}else if(hexDump){
				Readmemory(&buf, result.u, dumpLen, MM_RESTORE);
				char* hDump = HexDump(buf, dumpLen);
				if(hasComment>0){
					logmsg(1,"%s.0x%x  %s = %s  (%s)",Entries[i].dllName , Entries[i].rvaBpx , Entries[i].LogExp, hDump,comment );
				}else{
					logmsg(1,"%s.0x%x  %s = %s ",Entries[i].dllName , Entries[i].rvaBpx , Entries[i].LogExp, hDump);
				}
				free(hDump);
			}else{ //normal numeric
				if(hasComment>0){
					logmsg(1,"%s.0x%x  %s = %x  (%s)",Entries[i].dllName , Entries[i].rvaBpx , Entries[i].LogExp, result.u ,comment );
				}else{
					logmsg(1,"%s.0x%x  %s = %x ",Entries[i].dllName , Entries[i].rvaBpx , Entries[i].LogExp, result.u );
				}
			}

		}
		else{
			logmsg(1,"Results failed");
		}
	}catch(...){
		//MessageBox(0,"Caught exception","",0);
		logmsg(1,"Caught Exception: %s.0x%x  %s",Entries[i].dllName , Entries[i].rvaBpx , Entries[i].LogExp);
	}

}



VOID CALLBACK OnTracer_BreakPoint(HWND hwnd,UINT uMsg,UINT_PTR idEvent,DWORD dwTime){

	KillTimer((HWND)0, my_timer);
	if(active_bp_va == -1) return;

	bool shouldAbort = false;
    bool wasOneOfMine = false;

	for(int i=0;i<=numEntries;i++){
		if(active_bp_va == Entries[i].va){
			wasOneOfMine = true;
			Entries[i].hitcnt++;
			if(Entries[i].LogExp[0]!=0) HandleTracerBp(i); //an evaluate expression is set for this one
			if(abort_trace_index == i) shouldAbort = true; //process all (in case multiple per va) b4 abort
		}
	}

	active_bp_va = -1;

	if(shouldAbort){
		MessageBox(0,"Hit abort index breakpoint","",0);
		DialogBox( myHnd,MAKEINTRESOURCE(IDD_FORMVIEW),0,(DLGPROC)MyDlgProc);
	}
	else{
		if(wasOneOfMine) Sendshortcut(PM_MAIN, 0, WM_KEYDOWN, 0, 0, VK_F9); //continue were done
	}


}

extc void _export cdecl ODBG_Pluginmainloop(DEBUG_EVENT *debugevent){

	if( (int)debugevent == 0) return;

	if(debugevent->dwDebugEventCode == EXIT_PROCESS_DEBUG_EVENT){
		my_timer = SetTimer(0,0,600,(TIMERPROC)ShowFileLog_TimerProc); //required to use timer	
		return;
	}
	
	if(debugevent->dwDebugEventCode == CREATE_PROCESS_DEBUG_EVENT){
		m_hProcess = (int)debugevent->u.CreateProcessInfo.hProcess;
		return;
	}
	
	if(numEntries == -1) return;

	if(debugevent->dwDebugEventCode == EXCEPTION_DEBUG_EVENT){

		Broadcast(WM_USER_CHALL,0,0);
		Redrawdisassembler();

		if(debugevent->u.Exception.ExceptionRecord.ExceptionCode == EXCEPTION_BREAKPOINT){
			unsigned int adr = (int)debugevent->u.Exception.ExceptionRecord.ExceptionAddress;
			if(debug) logmsg(1,"adr=%x  Entries[0].va=%x  threadid=%x", adr, Entries[0].va , debugevent->dwThreadId );
			active_bp_va = adr;
			my_timer = SetTimer(0,0,600,(TIMERPROC)OnTracer_BreakPoint); //required to use timer to let ui catchup	
			return;
		}
	}


	if(debugevent->dwDebugEventCode != LOAD_DLL_DEBUG_EVENT) return;

	try{ 

		char dllName[200] = {0};
		char buf[500] = {0};

		GetModuleFileNameFromHeader( m_hProcess, 
									 debugevent->u.LoadDll.hFile, 
									 (int)debugevent->u.LoadDll.lpBaseOfDll , 
									 (char*)&dllName, 
									 199
									);
		
		if( (int)dllName==0 ){
			logmsg(1,"Error getting dllname @ base %x", debugevent->u.LoadDll.lpBaseOfDll);
			return;
		}

		strlower(dllName);
			
		if(debug) logmsg(1,"ModBpx: Dll %s loaded at base %x", dllName, debugevent->u.LoadDll.lpBaseOfDll);
			
		for(int i=0;i<=numEntries;i++){
			//if(stricmp(Entries[i].dllName , dllName) == 0 ) {
			if(strstr(dllName, Entries[i].dllName) != 0){ 
				int base = Entries[i].rvaBpx + (int)debugevent->u.LoadDll.lpBaseOfDll;
				if(debug) logmsg(1,"ModBpx: Setting bpx on %s.%x", dllName, base);
				Entries[i].va = base;
				Entries[i].enabled =true;
				Setbreakpoint(base,  TY_ACTIVE, 0);
			}
		}

	}catch(...){
		logmsg(1,"***** ModBpx: Caught Error in main loop ******");
	}



}





/*
try{
	int n = MyExpression(exp, debugevent->dwThreadId );

	if(asciiDump){	
		char buf[200];
		memset(buf,0,200);
		Readmemory(&buf, n, 199, MM_RESTORE);
		
		if(debug){
			sprintf(msg, "n=%x exp ='%s'", n, buf);
			MessageBox(0,msg,"in asciidump",0);
		}
		
		Addtolist(0,1,"%s.0x%x  %s = %s",Entries[i].dllName , Entries[i].rvaBpx , Entries[i].LogExp, buf );
	}else{ //normal numeric
		Addtolist(0,1,"%s.0x%x  %s = %x",Entries[i].dllName , Entries[i].rvaBpx , Entries[i].LogExp, n );
	}
}catch(...){
	sprintf(msg, "Error in extraction or expression exp=%s  eip=%x",Entries[i].LogExp, adr);
	MessageBox(0,msg,"",0);
}

			

int GetRegister(int i, int threadId=0){

	
	t_thread* t;
	t = Findthread(Getcputhreadid());
	CONTEXT context; 
	context.ContextFlags = CONTEXT_INTEGER | CONTEXT_CONTROL ;
	GetThreadContext(t->thread, &context);

	switch(i){
		case REG_EAX: return context.Eax;
		case REG_EBX: return context.Ebx;
		case REG_ECX: return context.Ecx;
		case REG_EDX: return context.Edx;
		case REG_EDI: return context.Edi;
		case REG_ESI: return context.Esi;
		case REG_ESP: return context.Esp;
		case REG_EBP: return context.Ebp;
		case RS_EIP: return context.Eip;
	}


	
	//t_thread *th = Findthread(threadId);
	//t_reg tr;
	//tr =  th->reg; 

//	if(i == RS_EIP) return (int)tr.ip;

 //   return (int)tr.r[i];
	

}

  //this is way to limited compared to the built in Expression() 
int MyExpression(char* exp, int threadId=0){
	
	char msg[100];
	int regIndex = -1;
	int ret =0;
	int ret2=0;

	if(debug) MessageBox(0,exp,"in myexpression",0);

	bool deref = exp[0] == '[' ? true : false;
	if(strstr(exp,"eax") != NULL) regIndex = REG_EAX;
	if(strstr(exp,"ebx") != NULL) regIndex = REG_EBX;
	if(strstr(exp,"ecx") != NULL) regIndex = REG_ECX;
	if(strstr(exp,"edx") != NULL) regIndex = REG_EDX;
	if(strstr(exp,"edi") != NULL) regIndex = REG_EDI;
	if(strstr(exp,"esi") != NULL) regIndex = REG_ESI;
	if(strstr(exp,"ebp") != NULL) regIndex = REG_EBP;
	if(strstr(exp,"esp") != NULL) regIndex = REG_ESP;
	if(strstr(exp,"eip") != NULL) regIndex = RS_EIP;

	if(debug){
		sprintf(msg, "eip=%x  exp ='%s' deref=%x regIndex=%x", GetRegister(RS_EIP,threadId), exp, (int)deref, regIndex);
		MessageBox(0,msg,"",0);
	}

	if(regIndex == -1) return 0xBADC0DE;

	ret = GetRegister(regIndex,threadId);
	
	if(debug){
		sprintf(msg, "regIndex=%x value=%x", regIndex, ret);
		MessageBox(0,msg,"",0);
	}

	if(deref){
		//extc ulong   cdecl Readmemory(void *buf,ulong addr,ulong size,int mode);
		Readmemory(&ret2, ret, 4, MM_RESTORE);
		return ret2;
	}

	return ret;

}

*/



