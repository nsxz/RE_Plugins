/*
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:    David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA
*/

typedef struct{
    int dwFlag;
    int cbSize;
    int lpData;
} cpyData;

#include <direct.h>

#if !defined (INVALID_FILE_ATTRIBUTES) 
#define INVALID_FILE_ATTRIBUTES ((DWORD)-1) 
#endif

typedef BOOL (WINAPI *lpEnumProcessModules)( HANDLE, HMODULE*, DWORD, LPDWORD);
lpEnumProcessModules EnumProcessModules=NULL;

typedef DWORD (WINAPI *lpGetModuleBaseName)(HANDLE,HMODULE,LPCSTR,DWORD);
lpGetModuleBaseName GetModuleBaseName=NULL;


//basically used to give us a function pointer with right prototype
//and 24 byte empty buffer inline which we assemble commands into in the
//hook proceedure. 
//#define BLOCK _asm int 3
#define ALLOC_THUNK(prototype) __declspec(naked) prototype { _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop}	   

ALLOC_THUNK( HINTERNET  __stdcall Real_HttpOpenRequest(HINTERNET hConnect,LPCWSTR lpszVerb,LPCWSTR lpszObjectName,LPCWSTR lpszVersion,LPCWSTR lpszReferrer,LPCWSTR FAR * lplpszAcceptTypes,DWORD dwFlags,DWORD dwContext));
ALLOC_THUNK( HINTERNET __stdcall Real_InternetConnect(HINTERNET hInternet,LPCWSTR lpszServerName,INTERNET_PORT nServerPort,	LPCWSTR lpszUserName OPTIONAL,							LPCWSTR lpszPassword OPTIONAL,DWORD dwService,DWORD dwFlags,	DWORD dwContext));
ALLOC_THUNK( BOOL __stdcall Real_InternetReadFile(HINTERNET hFile,LPVOID lpBuffer,DWORD dwNumberOfBytesToRead,LPDWORD lpdwNumberOfBytesRead));
ALLOC_THUNK( BOOL __stdcall Real_InternetCrackUrl(LPCTSTR lpszUrl, DWORD dwUrlLength,DWORD dwFlags,LPURL_COMPONENTS lpUrlComponents));
ALLOC_THUNK(BOOL __stdcall Real_HttpSendRequest(HINTERNET hRequest,LPCTSTR lpszHeaders,DWORD dwHeadersLength,LPVOID lpOptional, DWORD dwOptionalLength));

ALLOC_THUNK( HINTERNET  __stdcall Real_HttpOpenRequestA(HINTERNET hConnect,LPCSTR lpszVerb,LPCSTR lpszObjectName,LPCSTR lpszVersion,LPCSTR lpszReferrer,LPCSTR FAR * lplpszAcceptTypes,DWORD dwFlags,DWORD dwContext));
ALLOC_THUNK( HINTERNET __stdcall Real_InternetConnectA(HINTERNET hInternet,LPCSTR lpszServerName,INTERNET_PORT nServerPort,	LPCSTR lpszUserName OPTIONAL,							LPCSTR lpszPassword OPTIONAL,DWORD dwService,DWORD dwFlags,	DWORD dwContext));
ALLOC_THUNK( BOOL __stdcall Real_InternetCrackUrlA(LPCTSTR lpszUrl, DWORD dwUrlLength,DWORD dwFlags,LPURL_COMPONENTS lpUrlComponents));
ALLOC_THUNK(BOOL __stdcall Real_HttpSendRequestA(HINTERNET hRequest,LPCTSTR lpszHeaders,DWORD dwHeadersLength,LPVOID lpOptional, DWORD dwOptionalLength));

/*
ALLOC_THUNK( int	  __stdcall Real_URLDownloadToFileA(int a0,char* a1, char* a2, DWORD a3, int a4));
ALLOC_THUNK( int	  __stdcall Real_URLDownloadToCacheFile(int a0,char* a1, char* a2, DWORD a3, DWORD a4, int a5));
*/

void msg(char);
void LogAPI(const char*, ...);

extern void WriteToFile(char*,int, int, int);

bool pidLogged=false;
bool Warned=false;
HWND hServer=0;
int DumpAt=0;

void log_proc_name(void){

   HMODULE hMod;
   DWORD cbNeeded;
   char buf[255]={0};

  // return;

   if(pidLogged) return;

   HANDLE hProcess = OpenProcess( PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, GetCurrentProcessId() );
   
   if (hProcess){

		//_asm int 3

		if(EnumProcessModules==NULL){
			HMODULE hLib = LoadLibrary(L"psapi.dll");
			if(hLib){
				EnumProcessModules = (lpEnumProcessModules) GetProcAddress(hLib, "EnumProcessModules");
				GetModuleBaseName = (lpGetModuleBaseName) GetProcAddress(hLib, "GetModuleBaseNameA");
			}
		}

		if(EnumProcessModules!=NULL && GetModuleBaseName!=NULL){
			   if ( EnumProcessModules( hProcess, &hMod, sizeof(hMod), &cbNeeded) ){
				   GetModuleBaseName( hProcess, hMod, (char*)buf, 255 );
				   pidLogged=true;
				   LogAPI(" **** ProcID: %d = %s ****", GetCurrentProcessId(), buf);	
			   }
		}

   }

}






void FindVBWindow(){
	char *vbIDEClassName = "ThunderFormDC" ;
	char *vbEXEClassName = "ThunderRT6FormDC" ;
	char *vbWindowCaption = "ApiLogger" ;

	hServer = FindWindowA( vbIDEClassName, vbWindowCaption );
	if(hServer==0) hServer = FindWindowA( vbEXEClassName, vbWindowCaption );
	

} 

void msg(char *Buffer){
  
  char buf2[255]={0};

  //need to disable file writes if msg starts with * (appinit+dllmain shit)
  if(Buffer[0]!='*'){
	WriteToFile(Buffer,strlen(Buffer),2,1);
	WriteToFile("\r\n",2,2,1);
  }

  if(hServer==0 || !IsWindow(hServer) ) FindVBWindow();
  if(hServer==0) return;

  cpyData cpStructData;
  
  cpStructData.cbSize = strlen(Buffer) ;
  cpStructData.lpData = (int)Buffer;
  cpStructData.dwFlag = 3;
  
  SendMessage(hServer, WM_COPYDATA, 0,(LPARAM)&cpStructData);

} 

void LogAPI(const char *format, ...)
{
	DWORD dwErr = GetLastError();
		
	if(format){
		char buf[1024]; 
		va_list args; 
		va_start(args,format); 
		try{
 			 _vsnprintf(buf,1024,format,args);
			 msg(buf);
		}
		catch(...){}
	}

	SetLastError(dwErr);
}


__declspec(naked) int CalledFrom(){ 
	
	_asm{
			 mov eax, [ebp+4]  //return address of parent function (were nekkid)
			 ret
	}
	
}

 

void hexdump(unsigned char* str, int len){
	
	char asc[19];
	int aspot=0;
    const hexline_length = 3*16+4;
	
	char *nl="\n";
	char *tmp = (char*)malloc(50);
	
	//if(nohex) return;

	msg(nl);

	for(int i=0;i< len;i++){

		sprintf(tmp, "%02x ", str[i]);
		msg(tmp);
		
		if( (int)str[i]>20 && (int)str[i] < 123 ) asc[aspot] = str[i];
		 else asc[aspot] = 0x2e;

		aspot++;
		if(aspot%16==0){
			asc[aspot]=0x00;
			sprintf(tmp,"    %s\n", asc);
			msg(tmp);
			aspot=0;
		}

	}

	if(aspot%16!=0){//print last ascii segment if not full line
		int spacer = hexline_length - (aspot*3);
		while(spacer--)	msg(" ");	
		asc[aspot]=0x00;
		sprintf(tmp, "%s\n",asc);
		msg(tmp);
	}
	
	msg(nl);
	free(tmp);


}