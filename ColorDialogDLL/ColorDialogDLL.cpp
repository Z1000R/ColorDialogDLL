#include "ColorDialogDLL.h"
#include "ColorDropDownList.h"
#include "stringUtil.h"
#include "resource.h"

#if defined _M_IX86
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_IA64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='ia64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif

extern HINSTANCE hInst_;

using namespace std;

COLORDIALOGDLL_API COLORREF WINAPI showColorComboDialog(const VARIANT& vTitle, const VARIANT& vMessage, LPSAFEARRAY* ppsa, const LONG lDefaultIndex = 0)
{
	if (!Z1000R::isStringV(vTitle) || !Z1000R::isStringV(vMessage))
		return -2;

	wstring wsTitle = Z1000R::convVstr2Wstr(vTitle);
	wstring wsMessage = Z1000R::convVstr2Wstr(vMessage);

	if (!ppsa)
		return -2;

	if ((*ppsa)->cDims != 1)
		return -2;

	VARTYPE vt;
	SafeArrayGetVartype(*ppsa, &vt);
	if (vt != VT_I4 && vt != VT_UI4)
		return -2;

	Z1000R::ColorDropDownListDialogInfo cddldi
	{
		wsTitle.c_str(),
		wsMessage.c_str(),
		ppsa,
		lDefaultIndex
	};

	INT_PTR result = DialogBoxParamW(hInst_, MAKEINTRESOURCE(IDD_DDLDIALOG), nullptr, Z1000R::DialogProc, (LPARAM)&cddldi);

	return (COLORREF)result;
}