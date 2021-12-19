#include "framework.h"
#include "resource.h"
#include "ColorDropDownList.h"

using namespace std;

namespace Z1000R
{
	// 設定する色
	vector<COLORREF> vColoList;

	void initDropDownList(HWND);

	INT_PTR CALLBACK DialogProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
	{
		static HWND hDropDownList;

		switch (msg)
		{
		case WM_COMMAND:
		{
			switch (LOWORD(wParam))
			{
			case IDOK:
			{
				LONG lResult = (LONG)SendMessageW(hDropDownList, CB_GETCURSEL, 0, 0);
				COLORREF cr;
				if (lResult == -1)
					cr = -1;
				else
					cr = vColoList[lResult];
				EndDialog(hDlg, (INT_PTR)cr);
				return (INT_PTR)TRUE;
			}	// end of case IDOK:
			case IDCANCEL:
			{
				EndDialog(hDlg, -1);
				return (INT_PTR)TRUE;
			}	// end of case IDCANCEL:
			}	// end of switch (LOWORD(wParam))
			break;
		}	// end of case WM_COMMAND:
		case WM_INITDIALOG:
		{
			INITCOMMONCONTROLSEX ic;
			ic.dwICC = ICC_COOL_CLASSES;
			ic.dwSize = sizeof(INITCOMMONCONTROLSEX);
			InitCommonControlsEx(&ic);

			hDropDownList = GetDlgItem(hDlg, IDC_DDLCOLOR);
			initDropDownList(hDropDownList);

			Z1000R::ColorDropDownListDialogInfo* pCddldi = (ColorDropDownListDialogInfo*)lParam;
			SetWindowTextW(hDlg, pCddldi->pwszTitle);
			HWND hMesasge = GetDlgItem(hDlg, IDC_MESSAGE);
			SetWindowTextW(hMesasge, pCddldi->pwszMessage);

			LONG lLBound, lUBound;
			SafeArrayGetLBound(*pCddldi->ppsa, 1, &lLBound);
			SafeArrayGetUBound(*pCddldi->ppsa, 1, &lUBound);

			LONG lIndex{ 0 };
			vColoList.resize(lUBound - lLBound + 1);
			for (LONG i = lLBound; i <= lUBound; ++i)
			{
				// 色項目の追加
				// COLORREF値は、vColoListで管理するので、ここでは空文字列を設定し、
				// WM_DRAWITEMでカラーコードを書き込む
				SendMessageW(hDropDownList, CB_ADDSTRING, 0, (LPARAM)L"");
				SafeArrayGetElement(*pCddldi->ppsa, &i, &vColoList[lIndex]);
				++lIndex;
			}
			// 初期選択項目
			// wParam
			//  Specifies the zero-based index of the string to select.
			//  If this parameter is -1, any current selection in the list is removed and the edit control is cleared.
			// lParam
			//  This parameter is not used.
			SendMessage(hDropDownList, CB_SETCURSEL, pCddldi->lDefaultIndex, 0);

			return (INT_PTR)TRUE;
		}	// end of case WM_INITDIALOG:
		case WM_DRAWITEM:
		{
			LPDRAWITEMSTRUCT lpDis = (LPDRAWITEMSTRUCT)lParam;

			if (lpDis->CtlID != IDC_DDLCOLOR)
				return (INT_PTR)FALSE;

			// 背景塗り潰し
			DWORD dwBackColor = GetSysColor(COLOR_WINDOW);
			HBRUSH hbrBack = CreateSolidBrush(dwBackColor);
			FillRect(lpDis->hDC, &lpDis->rcItem, hbrBack);
			DeleteObject(hbrBack);

			// 項目が設定されていなければ、処理を抜ける
			if (lpDis->itemID == -1)
				return (INT_PTR)TRUE;

			// 色塗り潰し領域設定
			// 右側にカラーコードを表示させるための余白を作る
			RECT rcColor{ lpDis->rcItem };
			rcColor.top += 2;
			rcColor.bottom -= 2;
			rcColor.left += 2;
			rcColor.right = (rcColor.bottom - rcColor.top) * 3;
			HBRUSH hbrColor = CreateSolidBrush(vColoList[lpDis->itemID]);
			FillRect(lpDis->hDC, &rcColor, hbrColor);
			DeleteObject(hbrColor);

			// フォーカス項目
			// ODS_SELECTED    0x0001
			// ODS_FOCUS       0x0010
			UINT uiState = ODS_FOCUS | ODS_SELECTED;

			// ドロップダウン直後は
			//	ODS_SELECTED	がON
			//	ODS_FOCUS		がOFF
			//	となっているため、いずれかのフラグがONであれば、フォーカス表示させる
			DWORD dwTextColor;
			if ((lpDis->itemState & uiState))
			{
				// 塗り潰した色の部分を避けて、右側のみをフォーカス表示（反転）させるため、left をオフセット
				RECT rcFocus{ lpDis->rcItem };
				rcFocus.left = (lpDis->rcItem.bottom - lpDis->rcItem.top) * 3 - 5;
				HBRUSH hbrFocus = GetSysColorBrush(COLOR_HIGHLIGHT);
				FillRect(lpDis->hDC, &rcFocus, hbrFocus);

				dwTextColor = GetSysColor(COLOR_HIGHLIGHTTEXT);
			}
			else
				dwTextColor = GetSysColor(COLOR_WINDOWTEXT);

			// カラーコード
			wchar_t szColor[16];
			wsprintf(szColor, L"%02X%02X%02X",
				GetRValue(vColoList[lpDis->itemID]),
				GetGValue(vColoList[lpDis->itemID]),
				GetBValue(vColoList[lpDis->itemID]));
			SetTextColor(lpDis->hDC, dwTextColor);
			SetBkMode(lpDis->hDC, TRANSPARENT);
			// 塗り潰した色の右側に表示するため、left をオフセットした位置に書き込む
			RECT rcText{ lpDis->rcItem };
			rcText.left = (lpDis->rcItem.bottom - lpDis->rcItem.top) * 3;
			DrawText(lpDis->hDC, szColor, -1, &rcText, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

			return (INT_PTR)TRUE;
		}	// end of case WM_DRAWITEM:
		default:
			break;
		}	// end of switch (msg)

		return (INT_PTR)FALSE;
	}	// end of INT_PTR CALLBACK DialogProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)

	void initDropDownList(HWND hDropDownList)
	{
		LONG_PTR lpStyle = GetWindowLongPtrW(hDropDownList, GWL_STYLE);
		//lpStyle &= ~CBS_DROPDOWN;		// 0x0002L
		lpStyle |= CBS_DROPDOWNLIST;	// 0x0003L
		lpStyle &= ~CBS_OWNERDRAWVARIABLE;
		lpStyle |= CBS_OWNERDRAWFIXED;
		lpStyle &= ~CBS_SORT;
		lpStyle |= CBS_HASSTRINGS;
		SetWindowLongPtrW(hDropDownList, GWL_STYLE, lpStyle);
	}
}	// end of namespace Z1000R