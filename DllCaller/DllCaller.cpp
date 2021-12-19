#include <SDKDDKVer.h>
#include <windows.h>

#include <iostream>
#include <string>
#include <vector>

#if defined _M_IX86
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_IA64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='ia64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif

using namespace std;

vector<COLORREF> vColoList
{
	0xFFFFFF,
	0xC0C0C0,
	0x808080,
	0x000000,
	0x0000FF,
	0x000080,
	0x00FFFF,
	0x008080,
	0x00FF00,
	0x008000,
	0xFFFF00,
	0x808000,
	0xFF0000,
	0x800000,
	0xFF00FF,
	0x800080
};

using pFuncDlg = COLORREF(*)(VARIANT, VARIANT, LPSAFEARRAY*, LONG);

int main()
{
#ifdef _DEBUG
	HMODULE hModule = LoadLibraryW(L"ColorDialogDLL.dll");
#else
	HMODULE hModule = LoadLibraryW(L"ColorDialogDLL.dll");
#endif

	if (!hModule)
	{
		wcout << L"LoadLibrary fail.\n";
		return 0;
	}

	pFuncDlg pfDlg = (pFuncDlg)GetProcAddress(hModule, "showColorCombeDialog");
	if (!pfDlg)
	{
		wcout << L"LoadLibrary fail.\n";
		FreeLibrary(hModule);
		return 0;
	}

	VARIANT vTitle;
	VariantInit(&vTitle);
	vTitle.vt = VT_BSTR;
	vTitle.bstrVal = SysAllocString(L"Request From DllCaller");

	VARIANT vMessage;
	VariantInit(&vMessage);
	vMessage.vt = VT_BSTR;
	vMessage.bstrVal = SysAllocString(L"Select color.");

	//typedef struct tagSAFEARRAYBOUND
	//{
	//    ULONG cElements;
	//    LONG  lLbound;
	//} SAFEARRAYBOUND, * LPSAFEARRAYBOUND;
	SAFEARRAYBOUND rgs{ vColoList.size(), 0};
	LPSAFEARRAY psa = SafeArrayCreate(VT_UI4, 1u, &rgs);

	if (!psa)
	{
		wchar_t wszMessage[64];
		wsprintfW(wszMessage, L"[0x%08X] SafeArrayCreate fail.\n", GetLastError());
		wcout << wszMessage;

		VariantClear(&vTitle);
		VariantClear(&vMessage);

		return 0;
	}

	LONG lLBound{ 0 };
	SafeArrayGetLBound(psa, 1, &lLBound);
	LONG lUBound{ 0 };
	SafeArrayGetUBound(psa, 1, &lUBound);

	LONG lIndex{ 0 };
	for (LONG i = lLBound; i <= lUBound; ++i)
	{
		SafeArrayPutElement(psa, &i, &vColoList[lIndex]);
		++lIndex;
	}

	COLORREF cr = pfDlg(vTitle, vMessage, &psa, 3);

	VariantClear(&vTitle);
	VariantClear(&vMessage);
	SafeArrayDestroy(psa);

	FreeLibrary(hModule);

	wchar_t wsColor[16];
	if (!(cr & 0x80000000))
		wsprintfW(wsColor, L"%02X%02X%02X\n", GetRValue(cr), GetGValue(cr), GetBValue(cr));
	else if (cr == -1)
		wsprintfW(wsColor, L"%s\n", L"Cancel selected.");
	else
		wsprintfW(wsColor, L"%s", L"An error has occurred.");

	wcout << wsColor;
}