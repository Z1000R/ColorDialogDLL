#include "stringUtil.h"

using namespace std;

namespace Z1000R
{
	BOOL isStringV(const VARIANT v)
	{
		return (v.vt == VT_BSTR) || (v.vt == (VT_BSTR | VT_BYREF));
	}

	wstring	convVstr2Wstr(const VARIANT v)
	{
		std::wstring ws;

		if (v.vt == VT_BSTR)
			ws = wstring(v.bstrVal, SysStringLen(v.bstrVal));
		else if (v.vt == (VT_BSTR | VT_BYREF))
			ws = wstring(*v.pbstrVal, SysStringLen(*v.pbstrVal));

		return ws;
	}
}