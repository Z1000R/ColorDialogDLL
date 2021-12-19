// dllmain.cpp : DLL アプリケーションのエントリ ポイントを定義します。
#include "framework.h"

HINSTANCE hInst_;

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD   ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    hInst_ = hModule;

    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

