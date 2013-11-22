
#include <iostream>
#include <stdio.h>
#include <tchar.h>
#include <windows.h>
#import <msxml6.dll> rename_namespace(_T("MSXML"))
using namespace std;

int main(int argc, char* argv[])
{
	HRESULT hr = CoInitialize(NULL); 
	if (SUCCEEDED(hr))
	{
		try
		{
			MSXML::IXMLDOMDocument2Ptr xmlDoc;
			hr = xmlDoc.CreateInstance(__uuidof(MSXML::DOMDocument60), NULL, CLSCTX_INPROC_SERVER);
			// TODO: if (FAILED(hr))...

			if (xmlDoc->load(_T("Result.xml")) != VARIANT_TRUE)
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


