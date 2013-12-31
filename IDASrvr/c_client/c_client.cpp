
#include <stdio.h>
#include <conio.h>
#include <string.h>
#include <windows.h>

const int WM_DISPLAY_TEXT = 3;

//this is a super quick and dirty demo of a C client..

typedef struct{
    int dwFlag;
    int cbSize;
    int lpData;
} cpyData;


HWND hServer;
WNDPROC oldProc;
cpyData CopyData;

int  m_debug = 0;
int  m_ServerHwnd = 0;
char* IDA_SERVER_NAME = "IDA_SERVER";
char m_msg[2020];
bool received_response = false;

#pragma warning (disable:4996)

HWND ReadReg(char* name){ //todo support multiple IDA idb's and not just last one opened...

	 char* baseKey = "Software\\VB and VBA Program Settings\\IPC\\Handles";
	 char tmp[20] = {0};
     unsigned long l = sizeof(tmp);
	 HKEY h;
	 
	 RegOpenKeyExA(HKEY_CURRENT_USER, baseKey, 0, KEY_READ, &h);
	 RegQueryValueExA(h, name, 0,0, (unsigned char*)tmp, &l);
	 RegCloseKey(h);

	 return (HWND)atoi(tmp);
}

int SendTextMessage(int hwnd, char *Buffer, int blen) 
{
		  char* nullString = "NULL";
		  if(blen==0){ //in case they are waiting on a message with data len..
				Buffer = nullString;
				blen=4;
		  }
		  if(m_debug){
			  printf("Trying to send message to %x size:%d\n", hwnd, blen);
			  printf(" msg> %s\n", Buffer);
		  }
		  if(Buffer[blen] != 0) Buffer[blen]=0; ;
		  cpyData cpStructData;  
		  cpStructData.cbSize = blen;
		  cpStructData.lpData = (int)Buffer;
		  cpStructData.dwFlag = 3;
		  return SendMessage((HWND)hwnd, WM_COPYDATA, (WPARAM)hwnd,(LPARAM)&cpStructData);  
}  

/*
bool SendTextMessage(char* name, char *Buffer, int blen) 
{
		  HWND h = ReadReg(name);
		  if(IsWindow(h) == 0){
			  printf("Could not find valid hwnd for server %s\n", name);
			  return false;
		  }
		  return SendTextMessage((int)h,Buffer,blen);
}  

bool SendIntMessage(char* name, int resp){
	char tmp[30]={0};
	sprintf(tmp, "%d", resp);
	if(m_debug) printf("SendIntMsg(%s, %s)", name, tmp);
	return SendTextMessage(name,tmp, sizeof(tmp));
}
*/

int SendIntMessage(int hwnd, int resp){
	char tmp[30]={0};
	sprintf(tmp, "%d", resp);
	if(m_debug) printf("SendIntMsg(%d, %s)", hwnd, tmp);
	return SendTextMessage(hwnd,tmp, sizeof(tmp));
} 

void HandleMsg(char* m){
	//Message Received from IDA do stuff here...
	printf("%s\n", m);
}

/* old method */		//these next 2 are a very simple implementation..SendMessage automatically blocks so this works...
int ReceiveInt(char* command, int hwnd){
	memset(m_msg,0,2020);
	received_response = false;
	SendTextMessage(hwnd,command,strlen(command)+1);
	return atoi(m_msg);
}

int NewReceiveInt(char* command, int hwnd){
	return SendTextMessage(hwnd,command,strlen(command)+1);
}

char* ReceiveText(char* command, int hwnd){
	memset(m_msg,0,2020);
	received_response = false;
	SendTextMessage(hwnd, command,strlen(command)+1);
	return m_msg;
}


LRESULT CALLBACK WindowProc(HWND hwnd,UINT uMsg,WPARAM wParam,LPARAM lParam){
		
		if( uMsg != WM_COPYDATA) return 0;
		if( lParam == 0) return 0;
		
		memcpy((void*)&CopyData, (void*)lParam, sizeof(cpyData));
    
		if( CopyData.dwFlag == 3 ){
			if( CopyData.cbSize >= sizeof(m_msg) ) CopyData.cbSize = sizeof(m_msg)-1;
			memcpy((void*)&m_msg[0], (void*)CopyData.lpData, CopyData.cbSize);
			if(m_debug)	printf("Message Received: %s \n", m_msg); 
			received_response = true;
		}
			
    return 0;
}

int main(int argc, char* argv[])
{
 
	system("cls");

	m_ServerHwnd = (int)CreateWindowA("EDIT","MESSAGE_WINDOW", 0, 0, 0, 0, 0, 0, 0, 0, 0);
	oldProc = (WNDPROC)SetWindowLongA((HWND)m_ServerHwnd, GWL_WNDPROC, (LONG)WindowProc);
	
	int IDA_HWND = (int)ReadReg(IDA_SERVER_NAME);
	if(!IsWindow((HWND)IDA_HWND)) IDA_HWND = 0;

	if( m_ServerHwnd == 0){
		printf("Could not create listener window to receive data on exiting...\n");
		printf("Press any key to exit..");
		getch();
		return 0;
	}

	if( IDA_HWND==0 ){
		printf("IDA Server window not found exiting...\n");
		printf("Press any key to exit..");
		getch();
		return 0;
	}
	
	printf("Listening for responses on hwnd: %d\n", m_ServerHwnd); 
	printf("Active IDA hwnd: %d\n\n", IDA_HWND);

	char buf[255];
	int ret = 0;
    char* sret = 0;

	sprintf(buf,"loadedfile:%d", m_ServerHwnd);
	sret = ReceiveText(buf,IDA_HWND); 	
    printf("Loaded IDB: %s\n", sret);

	sprintf(buf,"numfuncs:%d", m_ServerHwnd);
	ret = ReceiveInt(buf, IDA_HWND);
	printf("Function Count: %d\n", ret);

	ret = NewReceiveInt("numfuncs", IDA_HWND);
	printf("Function Count: %d  (new method)\n", ret);

	sprintf(buf,"funcstart:1:%d", m_ServerHwnd);
	int funcStart = ReceiveInt(buf, IDA_HWND);
	printf("First Func Start: 0x%x\n", funcStart);

	sprintf(buf,"funcend:1:%d", m_ServerHwnd);
	ret = ReceiveInt(buf, IDA_HWND);
	printf("First Func End: 0x%x\n", ret);

	sprintf(buf,"getasm:%d:%d", funcStart, m_ServerHwnd);
	sret = ReceiveText(buf,IDA_HWND); 	
    printf("First Func Disasm[0]: %s\n", sret);
	
	printf("Press any key to exit..");
	getch();
	
	return 0;

}