#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers

#include <Windows.h>
#include <tchar.h>
#include <Psapi.h>
#include <shlobj.h>

#include <vector>
#include <iostream>
#include <sstream>

#ifdef _DEBUG
#include <chrono>
#include <ctime>
#endif

#ifdef _DEBUG
// _CRTDBG_MAP_ALLOC in Preprocessor Definitions
#define new new(_NORMAL_BLOCK, __FILE__, __LINE__)
#endif

using namespace std;

#ifdef _DEBUG
#define DCODE__(s) s
#else
#define DCODE__(s)
#endif

#define LOG__(s) do \
{ \
    wstringstream buf; \
    buf << L"SAS: " << s << endl; \
    OutputDebugString(buf.str().c_str()); \
} while (false)

#define DLOG__(s) DCODE__(LOG__(s))

struct WindowInfo
{
    HWND hwnd;
    wstring file;
    DWORD pid;
    DCODE__
    (
        wstring title;
    );
};

// ---------------------------------------------------------------------
const int HOTKEY_ID_FORWARD = 12;
const int HOTKEY_ID_BACKWARDS = 13;
const int HOTKEY_ID_IS_RESTORE = 14;
const int HOTKEY_ID_EXIT_OR_PAUSE = 15;

IVirtualDesktopManager* vdm = nullptr;
vector<WindowInfo> winfoList;
bool isRestore = true;
bool isPause = false;

size_t nextWinIndex = 0;
// ---------------------------------------------------------------------

bool GetProcessFileName(DWORD pid, wstring& file)
{
    HANDLE ph = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
    if (ph == NULL)
    {
        return false;
    }

    bool rc = true;
    wchar_t filePath[MAX_PATH];
    DWORD len = GetProcessImageFileName(ph, filePath, MAX_PATH);
    if (len)
    {
        file.assign(filePath, len);
    }
    else
    {
        rc = false;
    }

    CloseHandle(ph);

    return rc;
}

BOOL CALLBACK EnumerateWindows(HWND hwnd, LPARAM lParam)
{
    if (!IsWindowVisible(hwnd))
    {
        return TRUE;
    }

    if (vdm)
    {
        BOOL isThisVDesktop;
        HRESULT hr = vdm->IsWindowOnCurrentVirtualDesktop(hwnd, &isThisVDesktop);
        if (FAILED(hr))
        {
            LOG__(L"ERROR: vdm->IsWindowOnCurrentVirtualDesktop: " << hex << L"0x" << hr);
            return TRUE;
        }

        if (!isThisVDesktop) { return TRUE; }
    }

    WINDOWINFO wi = { 0 };
    wi.cbSize = sizeof(wi);
    if (GetWindowInfo(hwnd, &wi))
    {
        if (wi.dwExStyle & WS_EX_NOACTIVATE ||
            wi.dwStyle & WS_POPUP ||
            wi.dwExStyle & WS_EX_TOOLWINDOW)
        {
            return TRUE;
        }
    }

    DCODE__
    (
        wchar_t title[MAX_PATH] = { 0 };
    GetWindowText(hwnd, title, MAX_PATH);
    );

    wstring& activeFile = *(wstring*)lParam;

    DWORD pid;
    DWORD tid = GetWindowThreadProcessId(hwnd, &pid);
    wstring ifile;

    if (!GetProcessFileName(pid, ifile))
    {
        return TRUE;
    }

    if (ifile != activeFile)
    {
        return TRUE;
    }

    WindowInfo winInfo = {
        hwnd, ifile, pid
#ifdef _DEBUG
        , title
#endif
    };

    winfoList.push_back(winInfo);

    return TRUE;
}

void HotkeySwitchWindows(bool isHoldingMenuKey, bool forward)
{
    HWND fghwnd = GetForegroundWindow();

    if (fghwnd == NULL)
    {
        LOG__(L"ERROR: GetForegroundWindow " << GetLastError());
        return;
    }

    DWORD pid;
    DWORD tid = GetWindowThreadProcessId(fghwnd, &pid);

    wstring procfile;
    if (!GetProcessFileName(pid, procfile))
    {
        DLOG__(L"ERROR: GetProcessFile: " << GetLastError());
        return;
    }

    if (!isHoldingMenuKey)
    {
        nextWinIndex = 0;
        winfoList.clear();

        EnumDesktopWindows(nullptr, EnumerateWindows, (LPARAM)&procfile);
    }

    DCODE__
    (
        LOG__(L"#----------------------------------- "
            << chrono::system_clock::to_time_t(chrono::system_clock::now()));
        for (vector<WindowInfo>::iterator iter = winfoList.begin(); iter != winfoList.end(); iter++)
        {
            WindowInfo& windowInfo = *iter;
            LOG__(windowInfo.file << ' ' << windowInfo.hwnd << ' ' << windowInfo.title);
        }
    );

    if (winfoList.size() < 2)
        return;

    while (true)
    {
        if (forward)
        {
            ++nextWinIndex;
            if (nextWinIndex >= winfoList.size())
            {
                nextWinIndex = 0;
            }
        }
        else
        {
            if (nextWinIndex == 0)
            {
                nextWinIndex = winfoList.size() - 1;
            }
            else
            {
                --nextWinIndex;
            }
        }

        WindowInfo& winfo = winfoList.at(nextWinIndex);

        if (IsIconic(winfo.hwnd))
        {
            if (!isRestore)
            {
                continue;
            }

            if (!ShowWindowAsync(winfo.hwnd, SW_RESTORE))
            {
                // can't restore elevated apps
                LOG__(L"ERROR: ShowWindowAsync: " << GetLastError()); // 5: ERROR_ACCESS_DENIED
                continue;
            }
        }
        SetForegroundWindow(winfo.hwnd);
        break;
    }
}

bool RegisterMyHotKey(int keyId, UINT modifier, UINT vk, LPCWSTR msg)
{
    if (RegisterHotKey(NULL, keyId, modifier, vk))
    {
        return true;
    }

    DWORD error = ::GetLastError();
    if (error == ERROR_HOTKEY_ALREADY_REGISTERED)
    {
        LOG__(L"Hotkey is already registered by other process. " << msg);
    }
    else
    {
        LOG__(L"ERROR: Failed to register hot key " << msg << L", error code : " << error);
    }

    return false;
}

bool RegiterAllMyHotKeys()
{
    // forward
    // alt + `(backtick)
    if (!RegisterMyHotKey(HOTKEY_ID_FORWARD, MOD_ALT, VK_OEM_3, L"alt + `"))
    {
        // the only necessary hotkey
        return false;
    }

    // backwards
    // alt + shift + `(backtick)
    if (!RegisterMyHotKey(HOTKEY_ID_BACKWARDS, MOD_ALT | MOD_SHIFT, VK_OEM_3, L"alt + shift + `"))
    {
        LOG__(L"backward disabled");
    }

    // toggle to restore minimized windows
    // ctrl + alt + `(backtick)
    if (!RegisterMyHotKey(HOTKEY_ID_IS_RESTORE, MOD_ALT | MOD_CONTROL, VK_OEM_3, L"ctrl + alt + `"))
    {
        LOG__(L"restore disabled");
    }

    return true;
}

void UnregisterAllMyHotKeys()
{
    UnregisterHotKey(NULL, HOTKEY_ID_FORWARD);
    UnregisterHotKey(NULL, HOTKEY_ID_BACKWARDS);
    UnregisterHotKey(NULL, HOTKEY_ID_IS_RESTORE);
}

bool InitApp()
{
    HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(hr))
    {
        LOG__(L"ERROR: CoInitializeEx: " << hex << L"0x" << hr);
    }
    else
    {
        hr = CoCreateInstance(CLSID_VirtualDesktopManager, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&vdm));
        if (FAILED(hr))
        {
            LOG__(L"ERROR: CoCreateInstance CLSID_VirtualDesktopManager: " << hex << L"0x" << hr);
        }
    }

    // exit or pause
    // ctrl + alt + shift + `(backtick)
    if (!RegisterMyHotKey(HOTKEY_ID_EXIT_OR_PAUSE, MOD_CONTROL | MOD_ALT | MOD_SHIFT, VK_OEM_3, L"ctrl + alt + shift + `"))
    {
        LOG__("exit disabled");
    }

    if (!RegiterAllMyHotKeys())
    {
        return false;
    }
    return true;
}

void EndApp()
{
    if (vdm)
    {
        vdm->Release();
        vdm = nullptr;
    }

    CoUninitialize();
}


void DoMyJob()
{
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0) != 0)
    {
        if (msg.message != WM_HOTKEY)
        {
            continue;
        }

        int hotkeyId = (int)msg.wParam;
        if (hotkeyId == HOTKEY_ID_EXIT_OR_PAUSE)
        {
            SHORT lshift = GetAsyncKeyState(VK_LSHIFT);
            SHORT rshift = GetAsyncKeyState(VK_RSHIFT);

            // exit
            if ((lshift & 0x8000) && (rshift == 0))
            {
                LOG__(L"See you later.");
                break;
            }
            // pause
            if ((lshift == 0) && (rshift & 0x8000))
            {
                isPause = !isPause;
                if (isPause)
                {
                    UnregisterAllMyHotKeys();
                }
                else
                {
                    RegiterAllMyHotKeys();
                    // could wait for other app to release the hot key
                }
            }
            continue;
        }

        if (hotkeyId == HOTKEY_ID_IS_RESTORE)
        {
            isRestore = !isRestore;
            continue;
        }

        if (hotkeyId != HOTKEY_ID_FORWARD && hotkeyId != HOTKEY_ID_BACKWARDS)
        {
            continue;
        }

        SHORT lmenu = GetAsyncKeyState(VK_LMENU);
        SHORT rmenu = GetAsyncKeyState(VK_RMENU);
        bool isHoldingMenuKey = true; // holding the menu key and hit the backtick

        if (lmenu & 0x01 || rmenu & 0x01)
        {
            isHoldingMenuKey = false;
        }

        // UIPI (User Interface Privilege Isolation)
        // can't get the keystate of some elevated apps
        //   chrome, firefox, edge: good; their windows have the same process id
        //   ref: https://developer.chrome.com/blog/inside-browser-part1/
        // 
        // Maybe todo
        //   ref: https://docs.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/user-account-control-only-elevate-uiaccess-applications-that-are-installed-in-secure-locations?WT.mc_id=DT-MVP-4038148
        // 
        // incorrect but functioning anyway
        if (lmenu == 0 && rmenu == 0)
        {
            isHoldingMenuKey = false;
        }

        HotkeySwitchWindows(isHoldingMenuKey, hotkeyId == HOTKEY_ID_FORWARD);
    }
}

int APIENTRY wWinMain(_In_ HINSTANCE /*hInstance*/,
    _In_opt_ HINSTANCE /*hPrevInstance*/,
    _In_ LPWSTR    /*lpCmdLine*/,
    _In_ int       /*nCmdShow*/)
{
    DCODE__
    (
        _CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF);
    );

    if (!InitApp())
    {
        return 1;
    }

    LOG__(L"HotKey installed");

    DoMyJob();

    EndApp();

    return 0;
}
