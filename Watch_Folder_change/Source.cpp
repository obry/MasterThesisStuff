#include <windows.h>
#include <tchar.h> 
#include <stdio.h>
#include <strsafe.h>

#include <stdlib.h>
#include <string>
#include <iostream>
#include <fstream>
#include <sstream>

#import <msxml6.dll> rename_namespace(_T("MSXML"))

#pragma comment(lib, "User32.lib")

using namespace std;

void DisplayErrorBox(LPTSTR lpszFunction);


int ScanXML(string LatestFolderName)
{
	HRESULT hr = CoInitialize(NULL); 
	cout <<"Input Varibale: "+LatestFolderName <<endl;
	if (SUCCEEDED(hr))
	{
		try
		{
			MSXML::IXMLDOMDocument2Ptr xmlDoc;
			hr = xmlDoc.CreateInstance(__uuidof(MSXML::DOMDocument60), NULL, CLSCTX_INPROC_SERVER);
			// TODO: if (FAILED(hr))...

			// Convert 'string LatestFolderName' to something that xmlDoc->load() is able to read.
			// I know that 'ffd.cFileName' where I start with this is already something like wchar_t
			// But it works now, and I don't know yet how to combine two wchar_t like I need below etc.
			// Who ever reads this. Forgive me :))
			string tempLatestFolderName = LatestFolderName+"\\Result.xml";
		        wstring widestr = wstring(tempLatestFolderName.begin(), tempLatestFolderName.end());
			const wchar_t * widecstr = widestr.c_str();
			if (xmlDoc->load(widecstr) != VARIANT_TRUE)
			{
				printf("Unable to load Result.xml\n");
			}
			else
			{
				printf("XML was successfully loaded\n");

				xmlDoc->setProperty("SelectionLanguage", "XPath");

				MSXML::IXMLDOMNodeListPtr PeakAreaPercent = xmlDoc->getElementsByTagName("AreaPercent"); //Pointer to List of Elements with specified name
				int NumberOfPeaks = PeakAreaPercent->length; //Number of Peaks should be the number of AreaPercent entrys in XML File.
				cout <<"NumberOfPeaks=" << NumberOfPeaks <<endl;

				float * AreaPercentValue =new float[NumberOfPeaks];

				MSXML::IXMLDOMElementPtr spElementTemp; //need this to access values in elements
				for (int i = 0; i < PeakAreaPercent->length; i++) 
				{
					spElementTemp = (MSXML::IXMLDOMElementPtr) PeakAreaPercent->item[i];
					// Get the text node with the element content. If present it will be the first child.
					MSXML::IXMLDOMTextPtr spText = spElementTemp->firstChild;
					AreaPercentValue[i]=float(spText->nodeValue);
					if (spText != NULL) 
					{
						cout << " Element content: " << _bstr_t(spText->nodeValue) << endl;
					}
				}
			}
		}
		catch (_com_error &e)
		{
			printf("ERROR: %ws\n", e.ErrorMessage());
		}
		CoUninitialize();
	}
	return 0;
}


int _tmain(int argc, TCHAR *argv[])
{
	WIN32_FIND_DATA ffd;
	LARGE_INTEGER filesize;
	TCHAR szDir[MAX_PATH];
	size_t length_of_arg;
	HANDLE hFind = INVALID_HANDLE_VALUE;
	DWORD dwError=0;
	FILETIME  time1;
	FILETIME  time2;

	// If the directory is not specified as a command-line argument,
	// print usage.

	if(argc != 2)
	{
		_tprintf(TEXT("\nUsage: %s <directory name>\n"), argv[0]);
		return (-1);
	}

	// Check that the input path plus 3 is not longer than MAX_PATH.
	// Three characters are for the "\*" plus NULL appended below.

	StringCchLength(argv[1], MAX_PATH, &length_of_arg);

	if (length_of_arg > (MAX_PATH - 3))
	{
		_tprintf(TEXT("\nDirectory path is too long.\n"));
		return (-1);
	}

	_tprintf(TEXT("\nTarget directory is %s\n\n"), argv[1]);

	// Prepare string for use with FindFile functions.  First, copy the
	// string to a buffer, then append '\*' to the directory name.

	StringCchCopy(szDir, MAX_PATH, argv[1]);
	StringCchCat(szDir, MAX_PATH, TEXT("\\*"));

	// Find the first file in the directory.

	hFind = FindFirstFile(szDir, &ffd);

	if (INVALID_HANDLE_VALUE == hFind) 
	{
		DisplayErrorBox(TEXT("FindFirstFile"));
		return dwError;
	} 

	// List all the files in the directory with some info about them.
	int CompareFileTimeResult;
	bool firstRun=TRUE;
	bool firstCycle=TRUE;
	//definitoins used in conversion from wide char output of 'ffd.cFileName' to narrow char array
	char ch[260];
	char DefChar = ' ';
	string LatestFolderName;
	//wchar_t * LatestFolderNameWCHAR;  //tried to use wchar_t directly without converting to sring. Haven't got it to work yet.
	while(1)
	{
		do
		{
			if (ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
			{
				time1 = ffd.ftCreationTime;
				if (firstRun) 
				{ 
					time2 = time1;
					cout << "first run"<<endl;
				}
				CompareFileTimeResult = CompareFileTime(&time1,&time2); //returns 1, 0 or -1 
				if (CompareFileTimeResult == 1)
				{
					time2= time1;
					if (!firstCycle)
					{
						cout << "First file time is later than second file time." <<endl;
						WideCharToMultiByte(CP_ACP,0,ffd.cFileName,-1, ch,260,&DefChar, NULL);
						LatestFolderName = string(ch);
						cout <<"Latest Foldername is: " << LatestFolderName <<endl;
						ScanXML(LatestFolderName);
						//LatestFolderNameWCHAR = ffd.cFileName;
					}
				}
				/*if (CompareFileTimeResult == 0)
				{
				cout << "First file time is equal to second file time." <<endl;
				}
				if (CompareFileTimeResult == -1)
				{
				cout << "First file time is earlier than second file time." <<endl;
				}
				*/
				firstRun = FALSE;
			}
			/*else
			{
			filesize.LowPart = ffd.nFileSizeLow;
			filesize.HighPart = ffd.nFileSizeHigh;
			_tprintf(TEXT("  %s   %ld bytes\n"), ffd.cFileName, filesize.QuadPart);
			}*/
		}
		while (FindNextFile(hFind, &ffd) != 0);
		firstCycle=FALSE;
		cout << "end of run" <<endl;
		Sleep(5000);
		hFind = FindFirstFile(szDir, &ffd);
	}
	dwError = GetLastError();
	if (dwError != ERROR_NO_MORE_FILES) 
	{
		DisplayErrorBox(TEXT("FindFirstFile"));
	}

	FindClose(hFind);
	return dwError;
}


void DisplayErrorBox(LPTSTR lpszFunction) 
{ 
	// Retrieve the system error message for the last-error code

	LPVOID lpMsgBuf;
	LPVOID lpDisplayBuf;
	DWORD dw = GetLastError(); 

	FormatMessage(
		FORMAT_MESSAGE_ALLOCATE_BUFFER | 
		FORMAT_MESSAGE_FROM_SYSTEM |
		FORMAT_MESSAGE_IGNORE_INSERTS,
		NULL,
		dw,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
		(LPTSTR) &lpMsgBuf,
		0, NULL );

	// Display the error message and clean up

	lpDisplayBuf = (LPVOID)LocalAlloc(LMEM_ZEROINIT, 
		(lstrlen((LPCTSTR)lpMsgBuf)+lstrlen((LPCTSTR)lpszFunction)+40)*sizeof(TCHAR)); 
	StringCchPrintf((LPTSTR)lpDisplayBuf, 
		LocalSize(lpDisplayBuf) / sizeof(TCHAR),
		TEXT("%s failed with error %d: %s"), 
		lpszFunction, dw, lpMsgBuf); 
	MessageBox(NULL, (LPCTSTR)lpDisplayBuf, TEXT("Error"), MB_OK); 

	LocalFree(lpMsgBuf);
	LocalFree(lpDisplayBuf);
}

