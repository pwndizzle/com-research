// Proof of concept showing COM method interaction using IDispatch in C++
// Adapted from the code found here - http://www.cplusplus.com/forum/windows/125996/
// A way simpler Powershell equivalent:
// ([System.Activator]::CreateInstance([Type]::GetTypeFromCLSID("C08AFD90-F2A1-11D1-8455-00A0C91F3880"))).Document.Application.ShellExecute("cmd","/c notepad")

#include "stdafx.h"
#include <windows.h>
#include <ole2.h>
#include <Propvarutil.h>
#include <comdef.h>
#include <mshtml.h>
#include <exdisp.h>

const CLSID CLSID_SBApplication = { 0xC08AFD90,0xF2A1,0x11D1,{ 0x84,0x55,0x00,0xA0,0xC9,0x1F,0x38,0x80 } }; // CLSID of ShellBrowserWindow
const IID   IID_SBApplication = { 0xd30c1661,0xcdaf,0x11d0,{ 0x8a,0x3e,0x00,0xc0,0x4f,0xc9,0xe2,0x6e } }; // IID of ShellBrowserWindow

using namespace std;

int main()
{
	DISPPARAMS NoArgs = { NULL,NULL,0,0 };
	IDispatch* pSBApp = NULL;
	IDispatch* pSBDocument = NULL;
	IDispatch* pSBApplication = NULL;
	DISPPARAMS DispParams;
	VARIANT CallArgs[2];
	VARIANT vResult;
	DISPID dispid;
	HRESULT hr;

	CoInitialize(NULL);
	hr = CoCreateInstance(CLSID_SBApplication, NULL, CLSCTX_LOCAL_SERVER, IID_SBApplication, (void**)&pSBApp);
	if (SUCCEEDED(hr))
	{
		OLECHAR* szDocument = (OLECHAR*)L"Document";
		hr = pSBApp->GetIDsOfNames(IID_NULL, &szDocument, 1, GetUserDefaultLCID(), &dispid);
		if (SUCCEEDED(hr))
		{
			VariantInit(&vResult); 
			hr = pSBApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &NoArgs, &vResult, NULL, NULL);
			if (SUCCEEDED(hr))
			{
				pSBDocument = vResult.pdispVal;
				OLECHAR* szApplication = (OLECHAR*)L"Application";
				hr = pSBDocument->GetIDsOfNames(IID_NULL, &szApplication, 1, GetUserDefaultLCID(), &dispid);
				if (SUCCEEDED(hr))
				{
					VariantInit(&vResult);
					hr = pSBDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &NoArgs, &vResult, NULL, NULL);
					if (SUCCEEDED(hr))
					{
						pSBApplication = vResult.pdispVal;
						OLECHAR* szShellexecute = (OLECHAR*)L"ShellExecute";
						hr = pSBApplication->GetIDsOfNames(IID_NULL, &szShellexecute, 1, GetUserDefaultLCID(), &dispid);
						if (SUCCEEDED(hr))
						{
							CallArgs[0].vt = VT_BSTR;
							CallArgs[0].bstrVal = SysAllocString(L"/c notepad");
							CallArgs[1].vt = VT_BSTR;
							CallArgs[1].bstrVal = SysAllocString(L"cmd");

							DISPID dispidNamed = DISPID_PROPERTYPUT;
							DispParams.rgvarg = CallArgs;
							DispParams.rgdispidNamedArgs = &dispidNamed;
							DispParams.cArgs = 2;
							DispParams.cNamedArgs = 0;
							hr = pSBApplication->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &DispParams, &vResult, NULL, NULL);
						}
					}
				}
			}
		}
		VariantInit(&vResult);
		hr = pSBApp->Invoke(0x0000012e, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &NoArgs, &vResult, NULL, NULL);
		pSBApp->Release();
	}
	CoUninitialize();
	return 0;
}
