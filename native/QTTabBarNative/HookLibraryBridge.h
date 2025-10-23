#pragma once

#include <windows.h>

#include <optional>

namespace qttabbar::hooks {

struct HookCallbacks {
    void(__stdcall* hookResult)(int hookId, int retcode, void* context) = nullptr;
    BOOL(__stdcall* newWindow)(PCIDLIST_ABSOLUTE pidl, void* context) = nullptr;
    void* context = nullptr;
};

class HookLibraryBridge {
public:
    static HookLibraryBridge& Instance();

    HRESULT Initialize(const HookCallbacks& callbacks, const wchar_t* libraryPath);
    void Shutdown();
    HRESULT InitShellBrowserHook(IUnknown* shellBrowser);

private:
    HookLibraryBridge() = default;

    static void __cdecl ForwardHookResult(int hookId, int retcode);
    static bool __cdecl ForwardNewWindow(LPCITEMIDLIST pidl);

    HookCallbacks callbacks_{};
    HMODULE module_ = nullptr;
};

}  // namespace qttabbar::hooks

