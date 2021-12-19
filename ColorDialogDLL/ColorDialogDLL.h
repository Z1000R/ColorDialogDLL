#pragma once
#include "framework.h"

#define COLORDIALOGDLL_API __declspec(dllexport) 

extern "C"
{
	COLORDIALOGDLL_API COLORREF WINAPI showColorComboDialog(const VARIANT& vTitle, const VARIANT& vMessage, LPSAFEARRAY* ppsa, const LONG lDefaultIndex);
}