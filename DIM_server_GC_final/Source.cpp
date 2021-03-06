#include <dis.hxx>
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

#pragma comment(lib, "User32.lib") // _bstr_t is in here.

using namespace std;

typedef struct{
	float farr[12]; // this is more than needed, but right Number of Peaks is returned by GetNumberOfPeaks()
	// Open question: How to deal with this in PVSS? There must be better ideas.
	// ==> TPC prefers single service for each datapointelement anyway.
}COMPLEXDATA;

void DisplayErrorBox(LPTSTR lpszFunction);

string from_variant(VARIANT& vt)
	//conversion from _bstr_t to std::string
{
	_bstr_t bs(vt);
	return std::string(static_cast<const char*>(bs));
}

int GetStreamNumber(string LatestFolderName)
	// returns the Name of the Stream (Stream 1, 2, 3 ... 10) that was used by the Gaschromatograph's Stream-Selector,
	// by reading "SampleName" from "Result.xml"
{
	HRESULT hr = CoInitialize(NULL); 
	if (SUCCEEDED(hr))
	{
		try
		{
			MSXML::IXMLDOMDocument2Ptr xmlDoc;
			hr = xmlDoc.CreateInstance(__uuidof(MSXML::DOMDocument60), NULL, CLSCTX_INPROC_SERVER);
			string tempLatestFolderName = LatestFolderName+"/Result.xml";
			wstring widestr = wstring(tempLatestFolderName.begin(), tempLatestFolderName.end());
			const wchar_t * widecstr = widestr.c_str();
			cout <<"tempLatestFolderName in GetStreamNumber() is =" <<tempLatestFolderName <<endl;
			if (xmlDoc->load(widecstr) != VARIANT_TRUE)
			{
				printf("Unable to load Result.xml in function GetStreamNumber()\n");
			}
			else
			{
				xmlDoc->setProperty("SelectionLanguage", "XPath");
				// Find out from which Stream of the Chromatograph's Streamselector the XML file was produced.
				MSXML::IXMLDOMNodeListPtr SampleName = xmlDoc->getElementsByTagName("SampleName"); //Pointer to List of Elements with specified name
				MSXML::IXMLDOMElementPtr spElementTemp = (MSXML::IXMLDOMElementPtr) SampleName->item[0];
				MSXML::IXMLDOMTextPtr spText = spElementTemp->firstChild;
				// spText->nodeValu is e.g. "Stream 1". Extract the integer Number of the Stream:
				string StreamName = from_variant(spText->nodeValue);
				int StreamNumber = 0;
				string firstString= "";
				stringstream ss(StreamName);
				ss  >> firstString >> StreamNumber;
				return StreamNumber;
			}
		}
		catch (_com_error &e)
		{
			printf("ERROR: %ws\n", e.ErrorMessage());
		}
		CoUninitialize();
	}
	return -1;
}



string GetIjnectionTime(string LatestFolderName)
	// returns the InjectionDateTime
	// by reading "SampleName" from "Result.xml"
{
	HRESULT hr = CoInitialize(NULL); 
	if (SUCCEEDED(hr))
	{
		try
		{
			MSXML::IXMLDOMDocument2Ptr xmlDoc;
			hr = xmlDoc.CreateInstance(__uuidof(MSXML::DOMDocument60), NULL, CLSCTX_INPROC_SERVER);
			string tempLatestFolderName = LatestFolderName+"/Result.xml";
			wstring widestr = wstring(tempLatestFolderName.begin(), tempLatestFolderName.end());
			const wchar_t * widecstr = widestr.c_str();
			if (xmlDoc->load(widecstr) != VARIANT_TRUE)
			{
				printf("Unable to load Result.xml in function GetIjnectionTime()\n");
			}
			else
			{
				xmlDoc->setProperty("SelectionLanguage", "XPath");
				// Find out from which Stream of the Chromatograph's Streamselector the XML file was produced.
				MSXML::IXMLDOMNodeListPtr SampleName = xmlDoc->getElementsByTagName("InjectionDateTime"); //Pointer to List of Elements with specified name
				MSXML::IXMLDOMElementPtr spElementTemp = (MSXML::IXMLDOMElementPtr) SampleName->item[0];
				MSXML::IXMLDOMTextPtr spText = spElementTemp->firstChild;
				// spText->nodeValu is e.g. "Stream 1". Extract the integer Number of the Stream:
				string InjectionTime = from_variant(spText->nodeValue);
				return InjectionTime;
			}
		}
		catch (_com_error &e)
		{
			printf("ERROR: %ws\n", e.ErrorMessage());
		}
		CoUninitialize();
	}
	return "-1";
}
int GetNumberOfPeaks(string LatestFolderName)
	// returns the Number of Peaks that was found in "Result.xml"
{
	HRESULT hr = CoInitialize(NULL); 
	if (SUCCEEDED(hr))
	{
		try
		{
			MSXML::IXMLDOMDocument2Ptr xmlDoc;
			hr = xmlDoc.CreateInstance(__uuidof(MSXML::DOMDocument60), NULL, CLSCTX_INPROC_SERVER);
			string tempLatestFolderName = LatestFolderName+"/Result.xml";
			wstring widestr = wstring(tempLatestFolderName.begin(), tempLatestFolderName.end());
			const wchar_t * widecstr = widestr.c_str();
			if (xmlDoc->load(widecstr) != VARIANT_TRUE)
			{
				printf("Unable to load Result.xml in function GetNumberOfPeaks()\n");
			}
			else
			{
				xmlDoc->setProperty("SelectionLanguage", "XPath");
				MSXML::IXMLDOMNodeListPtr PeakAreaPercent = xmlDoc->getElementsByTagName("AreaPercent"); //Pointer to List of Elements with specified name
				int NumberOfPeaks = PeakAreaPercent->length; //Number of Peaks should be equal to the number of AreaPercent entrys in XML File.
				return NumberOfPeaks;
			}
		}
		catch (_com_error &e)
		{
			printf("ERROR: %ws\n", e.ErrorMessage());
		}
		CoUninitialize();
	}
	return -1;
}

COMPLEXDATA GetPeakData(string LatestFolderName, _bstr_t ElementName)
	//return Info with label 'ElementName' in "Result.xml" for all Peaks
{
	COMPLEXDATA Peaks;
	for (int i=0;i<12;i++)
	{
		Peaks.farr[i]=-1; //fill with dummy data to be able to return something in any case.
	}
	HRESULT hr = CoInitialize(NULL); 
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
			string tempLatestFolderName = LatestFolderName+"/Result.xml";
			wstring widestr = wstring(tempLatestFolderName.begin(), tempLatestFolderName.end());
			const wchar_t * widecstr = widestr.c_str();
			if (xmlDoc->load(widecstr) != VARIANT_TRUE)
			{
				printf("Unable to load Result.xml in GetPeakData()\n");
			}
			else
			{
				printf("XML was successfully loaded\n");

				xmlDoc->setProperty("SelectionLanguage", "XPath");

				MSXML::IXMLDOMNodeListPtr PeakAreaPercent = xmlDoc->getElementsByTagName(ElementName); //Pointer to List of Elements with specified name
				int NumberOfPeaks = PeakAreaPercent->length; //Number of Peaks should be equal to the number of AreaPercent entrys in XML File.
				float * AreaPercentValue =new float[NumberOfPeaks];

				MSXML::IXMLDOMElementPtr spElementTemp; //need this to access values in elements
				for (int i = 0; i < PeakAreaPercent->length; i++) 
				{
					spElementTemp = (MSXML::IXMLDOMElementPtr) PeakAreaPercent->item[i];
					// Get the text node with the element content. If present it will be the first child.
					MSXML::IXMLDOMTextPtr spText = spElementTemp->firstChild;
					AreaPercentValue[i]=(float)wcstod(_bstr_t(spText->nodeValue),NULL); //using directly _variant_t Extractor 'operator float()' will produce strange numbers. wcstof() is undefined according to compiler.
					Peaks.farr[i] = AreaPercentValue[i];
					//if (spText != NULL) 
					//{
					//cout << " Element TESTEST content: " << wcstod(_bstr_t(spText->nodeValue),NULL) << endl;
					//}
				}
				return Peaks;
			}
		}
		catch (_com_error &e)
		{
			printf("ERROR: %ws\n", e.ErrorMessage());
		}
		CoUninitialize();
	}
	return Peaks;
}


int _tmain(int argc, TCHAR *argv[])
	// Main function of this DIM Server.
	// Scans for new Subfolders in a Folder that needs to be specified at startup and calls above functions (e.g. GetPeakData()) to obtain Informatoin from "Result.xml" found in every subfolder.
	// (This is a modification of the code example found here: http://msdn.microsoft.com/en-us/library/aa365200(v=vs.85).aspx )	
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
	// string to a buffer, then append '/*' to the directory name.
	StringCchCopy(szDir, MAX_PATH, argv[1]);
	StringCchCat(szDir, MAX_PATH, TEXT("/*"));

	// Find the first file in the directory.
	hFind = FindFirstFile(szDir, &ffd);

	if (INVALID_HANDLE_VALUE == hFind) 
	{
		DisplayErrorBox(TEXT("FindFirstFile"));
		return dwError;
	} 

	// Declare / initialize variables that I use to find the latest subfolders in the monitored folder.
	int CompareFileTimeResult;
	bool firstRun=TRUE;
	bool firstCycle=TRUE;
	// Declare / initialize used in conversion from wide char output of e.g. 'ffd.cFileName' to narrow char array
	char ch[260];
	char DefChar = ' ';
	string LatestFolderName;
	// Convert userinput of Target Directory ( argv[1] ) to string in order to create full Path 'targetDirectory+"\\"+LatestFolderName' in the end.
	string targetDirectory;
	WideCharToMultiByte(CP_ACP,0,argv[1],-1, ch,260,&DefChar, NULL);
	targetDirectory = string(ch);	
	//initialize some return values
	int NumberOfPeaks;
	int StreamNumber;

	// DIM par starts here
	// create variables for DIM Services of TPC:
	float CO2_TPC=0;
	float Argon_TPC=0;
	float N2_TPC=0;
	float Water_TPC=0;
	// create Retention Time values:
	float RTCO2_min = 2.9f; //the "f" is to explicitly tell the compiler this is a float. I get "truncation from 'double' to 'float'" warning otherwise !?
	float RTCO2_max = 3.4f;
	float RTArgon_min = 5.9f;
	float RTArgon_max = 6.3f;
	// various stuff like Elementnames that will be called or temp variables:
	int activePeak = 0;
	_bstr_t RT = "RetTime";
	_bstr_t PeakArea = "AreaPercent";
	// create DIM-Services
	float * PeakAreasPointer = new float[];
	COMPLEXDATA PeakAreaData;
	COMPLEXDATA RTData;
	DimService Stream1("Stream1_PeakAreas","F:7",(void *)&PeakAreaData, sizeof(PeakAreaData));
	DimService tpcCO2Content("ALICE_GC.Actual.tpcCO2Content",CO2_TPC); 
	DimService tpcArgonContent("ALICE_GC.Actual.tpcArgonContent",Argon_TPC); 
	DimService tpcN2Content("ALICE_GC.Actual.tpcN2Content",N2_TPC); 
	DimService tpcWaterContent("ALICE_GC.Actual.tpcWaterContent",Water_TPC);
	DimServer::setDnsNode("alitpcdimdns"); //tpc dim dns node: 'alitpcdimdns'
	DimServer::autoStartOff();//Instructs the server to wait for a new DimServer::start(name) in order to register newly declared services. Lets me use two DimServer::start() with independent DIM_DNS_NODE for TRD and TPC.
	DimServer::start("ALICE_GC");

	string InjectionTimeAndDate_store1;
	string InjectionTimeAndDate_store2 = "this must not be empty" ;

	while(1)
	{
		do
		{
			if (ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) //only act on folder changes
			{
				time1 = ffd.ftLastWriteTime;
				if (firstRun) // only after startup of the program and only for the first folder found ...
				{ 
					time2 = time1;
					firstRun = FALSE;
				}
				CompareFileTimeResult = CompareFileTime(&time1,&time2); //returns 1, 0 or -1
				if ( CompareFileTimeResult == 1 ) // that is: First folder time is later than second folder time.
				{
					time1 = ffd.ftLastWriteTime;
					time2 = time1;
					if (!firstCycle)
					{
						StreamNumber = -1; //avoid that DIM services will be updated if no file could be read.
						WideCharToMultiByte(CP_ACP,0,ffd.cFileName,-1, ch,260,&DefChar, NULL);
						if ( string(ch) == "." )
						{
							cout << "WIN32_FIND_CHAR is seeing the parent folder change." <<endl;
						}
						else 
						{
							cout << "string(ch) as used = " << string(ch) << endl;
							LatestFolderName = targetDirectory+"/"+string(ch);
							cout<< "sleep 20 s. Count down from 10 ;)" <<endl;
							for (int i = 0; i<10;i++)
							{
								Sleep(2000);
								cout << 10-i <<endl;
							}
							InjectionTimeAndDate_store1 = GetIjnectionTime(LatestFolderName);
							cout << "InjectionTimeAndDate_store1: " << InjectionTimeAndDate_store1 <<endl;
							NumberOfPeaks = GetNumberOfPeaks(LatestFolderName);
							StreamNumber = GetStreamNumber(LatestFolderName);
							PeakAreaData = GetPeakData(LatestFolderName, PeakArea);
							RTData = GetPeakData(LatestFolderName, RT);
							if (InjectionTimeAndDate_store1 != InjectionTimeAndDate_store2) //only update DIM Services if xml file that was read is indeed a new file.
							{
								InjectionTimeAndDate_store2 = InjectionTimeAndDate_store1;
								for (int i=0;i<NumberOfPeaks;i++)
								{
									if ( StreamNumber == 1 ) // StreamNumber 1 is TPC GAS ==> update DIM Services for TPC
									{
										// decide based on Retention Time which peak is which and update DIM services accordingly:
										if ( RTCO2_min<RTData.farr[i] && RTData.farr[i]<RTCO2_max ) CO2_TPC = PeakAreaData.farr[i], cout << "CO2 candidate: " << CO2_TPC <<endl , tpcCO2Content.updateService();
										if ( RTArgon_min<RTData.farr[i] && RTData.farr[i]<RTArgon_max ) Argon_TPC = PeakAreaData.farr[i], cout << "Argon candidate: " << Argon_TPC <<endl , tpcArgonContent.updateService();
									}
									//cout << PeakArea <<" as COMPLEXDATA is : " << PeakAreaData.farr[i] << endl;
									//cout << RT <<" as COMPLEXDATA is : " << RTData.farr[i] << endl;
								}
								if (PeakAreaData.farr[0] >= 0 ) Stream1.updateService(); //only update Stream1 when it contains real data (that is no dummy data "-1").

							}
						}
					}
				}
			}
		}
		while (FindNextFile(hFind, &ffd) != 0);
		// <- end of do... while ....
		firstCycle=FALSE;
		Sleep(4000);
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
