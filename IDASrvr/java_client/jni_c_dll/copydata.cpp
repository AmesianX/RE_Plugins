
#include <stdio.h>
#include <string.h>
#include <windows.h>
#include "JNITest.h"
#include <jni.h>

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

int  IDA_HWND=0;
int  m_debug = 0;
int  m_ServerHwnd = 0;
char* IDA_SERVER_NAME = "IDA_SERVER";
char m_msg[2020];
bool received_response = false;
HANDLE hConOut = 0;

#pragma warning (disable:4996)

//enum colors{ mwhite=15, mgreen=10, mred=12, myellow=14, mblue=9, mpurple=5, mgrey=7, mdkgrey=8 };

void end_color(void){
	SetConsoleTextAttribute(hConOut,7); 
}

void start_color(/*enum colors c*/){
    SetConsoleTextAttribute(hConOut, 14);
}

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

bool SendTextMessage(int hwnd, char *Buffer, int blen) 
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
		  SendMessage((HWND)hwnd, WM_COPYDATA, (WPARAM)hwnd,(LPARAM)&cpStructData);  
		  return true;
}  

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

bool SendIntMessage(int hwnd, int resp){
	char tmp[30]={0};
	sprintf(tmp, "%d", resp);
	if(m_debug) printf("SendIntMsg(%d, %s)", hwnd, tmp);
	return SendTextMessage(hwnd,tmp, sizeof(tmp));
} 

//these next 2 are a very simple implementation..SendMessage automatically blocks so this works...
int ReceiveInt(char* command, int hwnd){
	memset(m_msg,0,2020);
	received_response = false;
	SendTextMessage(hwnd,command,strlen(command)+1);
	return atoi(m_msg);
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
			if(m_debug){
				start_color();
				printf("JNI: WindowProc Message Received: %s \n", m_msg); 
				end_color();
			}
			received_response = true;
		}
			
    return 0;
}

JNIEXPORT jint JNICALL Java_JNITest_SendCMDRecvInt(JNIEnv *env, jobject obj, jstring javaString)
{
	 char* cmd = (char*)env->GetStringUTFChars(javaString, 0);
     
	 start_color();
	 printf("JNI: SendCMDRecvInt(%s)\n", cmd);
	 end_color();

	 return ReceiveInt(cmd,IDA_HWND);
}

JNIEXPORT void JNICALL Java_JNITest_SendCMD(JNIEnv *env, jobject obj, jstring javaString)
{
	char* cmd = (char*)env->GetStringUTFChars(javaString, 0);

	start_color();
	printf("JNI: SendCMD(%s)\n", cmd); 
	end_color();

	SendTextMessage(IDA_HWND, cmd, strlen(cmd) );
	
	env->ReleaseStringUTFChars(javaString, cmd);
}


JNIEXPORT jstring JNICALL Java_JNITest_SendCMDRecvText(JNIEnv *env, jobject obj, jstring javaString)
{
	 char* cmd = (char*)env->GetStringUTFChars(javaString, 0);
     
	 start_color();
	 printf("JNI: SendCMDRecvText(%s)\n", cmd);
	 
	 char* resp = ReceiveText(cmd,IDA_HWND);

	 printf("JNI: received: %s\n", resp);
	 end_color();

	 jstring js_ret = env->NewStringUTF(resp);
	 env->ReleaseStringUTFChars(javaString, cmd);

	 return js_ret;


}

JNIEXPORT jint JNICALL Java_JNITest_InitHwnd(JNIEnv *env, jobject obj)
{
 
	hConOut = GetStdHandle( STD_OUTPUT_HANDLE );
	m_ServerHwnd = (int)CreateWindowA("EDIT","MESSAGE_WINDOW", 0, 0, 0, 0, 0, 0, 0, 0, 0);
	oldProc = (WNDPROC)SetWindowLongA((HWND)m_ServerHwnd, GWL_WNDPROC, (LONG)WindowProc);
	
	IDA_HWND = (int)ReadReg(IDA_SERVER_NAME);
	if(!IsWindow((HWND)IDA_HWND)) IDA_HWND = 0;

	start_color();

	if( m_ServerHwnd == 0){
		printf("JNI: Could not create listener window to receive data on\n");
		end_color();
		return 0;
	}

	if( IDA_HWND==0 ){
		printf("JNI: IDA Server window not found\n");
		end_color();
		return 0;
	}
	
	printf("JNI: Listening for responses on hwnd: %d\n", m_ServerHwnd); 
	printf("JNI: Active IDA hwnd: %d\n", IDA_HWND);
	end_color();

	return m_ServerHwnd;	

}