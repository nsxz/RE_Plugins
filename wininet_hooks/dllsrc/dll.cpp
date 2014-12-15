/*

	hooking library from old gpl work used, 

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

#define UNICODE

#include <windows.h>
#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <stdarg.h>
#include <wininet.h>



void InstallHooks(void);

extern "C" void __setargv(void);


#include "hooker.h"
#include "main.h"   //contains a bunch of library functions in it too..

bool Installed =false;
  

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{

    if(!Installed){
		 Installed=true;
		 InstallHooks();
	}

	return TRUE;
}

int file_exists(char* pth){
	FILE *fp = fopen(pth,"r");
	if(fp) {
		fclose(fp);
		return 1;
	}
	return 0;
}

int folder_exists(char* strFolderName)
{   
    if(GetFileAttributesA(strFolderName) != INVALID_FILE_ATTRIBUTES){
		if(file_exists(strFolderName)==0) return 1;
	}
	return 0;
}

int newfile(char* basedir){
	
	char tmp[255]={0};

	for(int i=0;i<5000;i++){
		sprintf(tmp, "%s\\%x.txt",basedir,i);
		if(file_exists(tmp)==0){
			strcpy(basedir,tmp);
			return 1;
		}
	}
	return 0;
}

			

void WriteToFile(char* buf,int leng, int which=0, int silent=0)
{
  FILE * pFile;
  char* nl = "\n\n";
  char buf2[255]={20};
  char pth[255]={0};

  if(!folder_exists("c:\\pages")) mkdir("c:\\pages");
  if(!folder_exists("c:\\posts")) mkdir("c:\\posts");

  switch(which){
	case 0: strcpy(pth,"c:\\pages");break;
	case 1: strcpy(pth,"c:\\posts");break;
    case 2: strcpy(pth,"c:\\wininet_log.txt");break;
  }

  if(which < 2){
	  if(newfile(pth)==0){
		if(silent==0) LogAPI("newfile returned 0 write failed?");
		return;
	  }
  }  
  
  if(silent==0) LogAPI("Writing to %s", pth);
  pFile = fopen ( pth , "ab" );
  fwrite(buf , 1 , leng , pFile );
  fclose (pFile);

}

void stripnulls(char* bin, int bin_len, char* bout, int bout_len){

	int max = bout_len;
    int j=0;

	//_asm int 3

	for(int i=0;i<max;i++){
		if(bin[i]==0 && bin[i+1]==0) break;
		if(bin[i]!=0){
			 bout[j] = bin[i];
			 j++;
		}
		i++;
	}

}

BOOL __stdcall My_InternetCrackUrlA(
       LPCTSTR lpszUrl,
       DWORD dwUrlLength,
       DWORD dwFlags,
	   LPURL_COMPONENTS lpUrlComponents
	   ){
	int mylen=dwUrlLength;
	
	log_proc_name();

	if(lpszUrl!=NULL){
	    if(mylen==0) mylen=strlen((char*)lpszUrl);
		LogAPI("%d.%x> InternetCrackUrlA: %s", GetCurrentProcessId(), CalledFrom(), lpszUrl);
	}

	return Real_InternetCrackUrlA(lpszUrl,dwUrlLength,dwFlags,lpUrlComponents);

}

BOOL __stdcall My_InternetCrackUrl(
       LPCTSTR lpszUrl,
       DWORD dwUrlLength,
       DWORD dwFlags,
	   LPURL_COMPONENTS lpUrlComponents
	   ){

	char buf[2000]={0};
	int mylen=dwUrlLength;

	log_proc_name();

	if(lpszUrl!=NULL){
	    if(mylen==0) mylen=strlen((char*)lpszUrl);
		stripnulls((char*)lpszUrl,mylen, &buf[0], 2000);
		LogAPI("%d.%x> InternetCrackUrlW: %s", GetCurrentProcessId(), CalledFrom(), buf);
	}

	return Real_InternetCrackUrl(lpszUrl,dwUrlLength,dwFlags,lpUrlComponents);

}


BOOL __stdcall My_HttpSendRequestA(
    HINTERNET hRequest,
    LPCTSTR lpszHeaders,
    DWORD dwHeadersLength,
    LPVOID lpOptional,
    DWORD dwOptionalLength
){
	
	log_proc_name();

	if(lpszHeaders!=NULL){
		LogAPI("%d.%x> HttpSendRequestA Headers=%s", GetCurrentProcessId(), CalledFrom(), lpszHeaders);
	}

	if(dwOptionalLength>0 && lpOptional!=NULL){
	     WriteToFile((char*)lpOptional,dwOptionalLength,1); 
		 LogAPI("%d.%x> HttpSendRequestA Post Data written to disk Len=%x", GetCurrentProcessId(), CalledFrom(), dwOptionalLength);
	}

	return Real_HttpSendRequestA(hRequest,lpszHeaders,dwHeadersLength,lpOptional,dwOptionalLength);


}

BOOL __stdcall My_HttpSendRequest(
    HINTERNET hRequest,
    LPCTSTR lpszHeaders,
    DWORD dwHeadersLength,
    LPVOID lpOptional,
    DWORD dwOptionalLength
){
	
	char buf[5000]={0};
	log_proc_name();

	if(lpszHeaders!=NULL){
		stripnulls((char*)lpszHeaders,dwHeadersLength, &buf[0], 5000);
		LogAPI("%d.%x> HttpSendRequest Headers=%s", GetCurrentProcessId(), CalledFrom(), buf);
	}

	if(dwOptionalLength>0 && lpOptional!=NULL){
	     WriteToFile((char*)lpOptional,dwOptionalLength,1); 
		 LogAPI("%d.%x> HttpSendRequestW Post Data written to disk Len=%x", GetCurrentProcessId(), CalledFrom(), dwOptionalLength);
	}

	return Real_HttpSendRequest(hRequest,lpszHeaders,dwHeadersLength,lpOptional,dwOptionalLength);


}


BOOL __stdcall My_InternetReadFile(
	HINTERNET hFile,
	LPVOID lpBuffer,
	DWORD dwNumberOfBytesToRead,
	LPDWORD lpdwNumberOfBytesRead
){

	char buf[1024]={0};
	int i=0, j=0;
	char* tmp = (char*)lpBuffer;

	log_proc_name();

	//_asm int 3

	BOOL ret = Real_InternetReadFile(hFile,lpBuffer, dwNumberOfBytesToRead,lpdwNumberOfBytesRead);
	
	if(ret){
		WriteToFile(tmp,dwNumberOfBytesToRead);
		LogAPI("%d.%x> InternetReadFile %x bytes logged to disk", GetCurrentProcessId(), CalledFrom(), dwNumberOfBytesToRead);
	}

	return ret;

}



HINTERNET  __stdcall My_HttpOpenRequestA(
									HINTERNET hConnect,
									LPCSTR lpszVerb,
									LPCSTR lpszObjectName,
									LPCSTR lpszVersion,
									LPCSTR lpszReferrer,
									LPCSTR FAR * lplpszAcceptTypes,
									DWORD dwFlags,
									DWORD dwContext
									){

	log_proc_name();

	if(lpszObjectName!=NULL){
		LogAPI("%d.%x> HttpOpenRequestA: %s", GetCurrentProcessId(), CalledFrom(), lpszObjectName);
	}

	return Real_HttpOpenRequestA(hConnect,lpszVerb,lpszObjectName,lpszVersion,lpszReferrer,lplpszAcceptTypes,dwFlags,dwContext);

}



HINTERNET  __stdcall My_HttpOpenRequest(
									HINTERNET hConnect,
									LPCWSTR lpszVerb,
									LPCWSTR lpszObjectName,
									LPCWSTR lpszVersion,
									LPCWSTR lpszReferrer,
									LPCWSTR FAR * lplpszAcceptTypes,
									DWORD dwFlags,
									DWORD dwContext
									){

	char buf[1500]={0};
	int i=0, j=0;
	char* tmp = (char*)lpszObjectName;
	log_proc_name();

	//_asm int 3
	
	while(1){
		if(j >= sizeof(buf)) break;
		if(tmp[i]==0 && tmp[i+1]==0) break;
		if(tmp[i]!=0){
			buf[j] = tmp[i];
			j++;
		}
		i++;
	}
	
	LogAPI("%d.%x> HttpOpenRequestW: %s", GetCurrentProcessId(), CalledFrom(), (char*)buf);

	return Real_HttpOpenRequest(hConnect,lpszVerb,lpszObjectName,lpszVersion,lpszReferrer,lplpszAcceptTypes,dwFlags,dwContext);

}

HINTERNET __stdcall My_InternetConnect(
						HINTERNET hInternet,
						LPCWSTR lpszServerName,
						INTERNET_PORT nServerPort,
						LPCWSTR lpszUserName OPTIONAL,
						LPCWSTR lpszPassword OPTIONAL,
						DWORD dwService,
						DWORD dwFlags,
						DWORD dwContext){

	char buf[1500]={0};
	int i=0, j=0;
	char* tmp = (char*)lpszServerName;
	log_proc_name();

	//_asm int 3
	
	while(1){
		if(j >= sizeof(buf)) break;
		if(tmp[i]==0 && tmp[i+1]==0) break;
		if(tmp[i]!=0){
			buf[j] = tmp[i];
			j++;
		}
		i++;
	}
	LogAPI("%d.%x> InternetConnectW: %s", GetCurrentProcessId(), CalledFrom(), buf);

	return Real_InternetConnect(hInternet,lpszServerName,nServerPort,lpszUserName,lpszPassword,dwService,dwFlags,dwContext);

}

HINTERNET __stdcall My_InternetConnectA(
						HINTERNET hInternet,
						LPCSTR lpszServerName,
						INTERNET_PORT nServerPort,
						LPCSTR lpszUserName OPTIONAL,
						LPCSTR lpszPassword OPTIONAL,
						DWORD dwService,
						DWORD dwFlags,
						DWORD dwContext){

	log_proc_name();

	if(lpszServerName!=NULL){
		LogAPI("%d.%x> InternetConnectA: %s", GetCurrentProcessId(), CalledFrom(), lpszServerName);
	}

	return Real_InternetConnectA(hInternet,lpszServerName,nServerPort,lpszUserName,lpszPassword,dwService,dwFlags,dwContext);

}


/*
int My_URLDownloadToFileA(int a0,char* a1, char* a2, DWORD a3, int a4)
{
	
	log_proc_name();
	LogAPI("%d.%x> URLDownloadToFile(%s)", GetCurrentProcessId(), CalledFrom(), a1);

    SOCKET ret = 0;
    try {
        ret = Real_URLDownloadToFileA(a0, a1, a2, a3, a4);
    }
	catch(...){	} 

    return ret;
}

//untested
int My_URLDownloadToCacheFile(int a0,char* a1, char* a2, DWORD a3, DWORD a4, int a5)
{
	
	log_proc_name();
	LogAPI("%d.%x> URLDownloadToCacheFile(%s)", GetCurrentProcessId(), CalledFrom(), a1);

    SOCKET ret = 0;
    try {
        ret = Real_URLDownloadToCacheFile(a0, a1, a2, a3, a4, a5);
    }
	catch(...){	} 

    return ret;
}
*/


void main(void){;/* an empty fx to export in case you want to add this dll to an import table manually */ }

void DoHook(void* real, void* hook, void* thunk, char* name){

	char err[400];

	if ( !InstallHook( real, hook, thunk) ){ //try to install the real hook here
		sprintf(err,"***** Install %s hook failed...Error: %s", name, &lastError);
		msg(err);
	} 

}

#define ADDHOOK(name) DoHook( name, My_##name, Real_##name, #name );	

void InstallHooks(void)
{

    //no fancyshit! sprintf or LogAPI are not allowed in the here (dllmain is a funny spot)
 	msg("***** Installing Hooks *****");	
 
	ADDHOOK(HttpOpenRequest);
	ADDHOOK(InternetConnect);
	ADDHOOK(InternetReadFile);
	ADDHOOK(InternetCrackUrl);
	ADDHOOK(HttpSendRequest);

    ADDHOOK(HttpOpenRequestA);
	ADDHOOK(InternetConnectA);
	ADDHOOK(InternetCrackUrlA);
	ADDHOOK(HttpSendRequestA);

    //ADDHOOK(URLDownloadToFileA);   //can not link to urlmon if you want to load via appinit_dlls!
	//ADDHOOK(URLDownloadToCacheFile);

}




