#pragma once
#include "framework.h"

namespace Z1000R
{
	struct ColorDropDownListDialogInfo
	{
		PCWCHAR pwszTitle;
		PCWCHAR pwszMessage;
		LPSAFEARRAY* ppsa;
		LONG lDefaultIndex;
	};

	INT_PTR CALLBACK DialogProc(HWND, UINT, WPARAM, LPARAM);
}
