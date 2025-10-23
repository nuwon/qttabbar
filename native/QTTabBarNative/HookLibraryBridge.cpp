#include "pch.h"

#include "HookLibraryBridge.h"

#include <shlwapi.h>

#include <mutex>

#pragma comment(lib, "Shlwapi.lib")

namespace qttabbar::hooks {

namespace {
struct CallbackStruct {
    void(__cdecl* hookResult)(int, int);
    bool(__cdecl* newWindow)(LPCITEMIDLIST);
};

using InitializeFn = int(__cdecl*)(CallbackStruct*);
using DisposeFn = int(__cdecl*)();
using InitShellBrowserHookFn = int(__cdecl*)(IShellBrowser*);

std::mutex g_mutex;
}  // namespace

HookLibraryBridge& HookLibraryBridge::Instance() {
    static HookLibraryBridge instance;
    return instance;
}

HRESULT HookLibraryBridge::Initialize(const HookCallbacks& callbacks, const wchar_t* libraryPath) {
    std::scoped_lock lock(g_mutex);

    if (module_ != nullptr) {
        callbacks_ = callbacks;
        return S_OK;
    }

    if (libraryPath == nullptr || !PathFileExistsW(libraryPath)) {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    module_ = ::LoadLibraryExW(libraryPath, nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS);
    if (module_ == nullptr) {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    auto initialize = reinterpret_cast<InitializeFn>(::GetProcAddress(module_, "Initialize"));
    if (initialize == nullptr) {
        Shutdown();
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    CallbackStruct callbacksNative{};
    callbacksNative.hookResult = &HookLibraryBridge::ForwardHookResult;
    callbacksNative.newWindow = &HookLibraryBridge::ForwardNewWindow;

    int result = initialize(&callbacksNative);
    if (result != 0) {
        Shutdown();
        return HRESULT_FROM_WIN32(ERROR_INVALID_FUNCTION);
    }

    callbacks_ = callbacks;
    return S_OK;
}

void HookLibraryBridge::Shutdown() {
    std::scoped_lock lock(g_mutex);
    if (module_ != nullptr) {
        if (auto dispose = reinterpret_cast<DisposeFn>(::GetProcAddress(module_, "Dispose"))) {
            dispose();
        }
        ::FreeLibrary(module_);
        module_ = nullptr;
    }
    callbacks_ = {};
}

HRESULT HookLibraryBridge::InitShellBrowserHook(IUnknown* shellBrowser) {
    if (module_ == nullptr || shellBrowser == nullptr) {
        return E_FAIL;
    }
    auto initHook = reinterpret_cast<InitShellBrowserHookFn>(::GetProcAddress(module_, "InitShellBrowserHook"));
    if (initHook == nullptr) {
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }
    CComPtr<IShellBrowser> spBrowser;
    HRESULT hr = shellBrowser->QueryInterface(IID_PPV_ARGS(&spBrowser));
    if (FAILED(hr)) {
        return hr;
    }
    return initHook(spBrowser) == 0 ? S_OK : E_FAIL;
}

void __cdecl HookLibraryBridge::ForwardHookResult(int hookId, int retcode) {
    HookLibraryBridge& bridge = HookLibraryBridge::Instance();
    if (bridge.callbacks_.hookResult != nullptr) {
        bridge.callbacks_.hookResult(hookId, retcode, bridge.callbacks_.context);
    }
}

bool __cdecl HookLibraryBridge::ForwardNewWindow(LPCITEMIDLIST pidl) {
    HookLibraryBridge& bridge = HookLibraryBridge::Instance();
    if (bridge.callbacks_.newWindow != nullptr) {
        return bridge.callbacks_.newWindow(pidl, bridge.callbacks_.context) != FALSE;
    }
    return false;
}

}  // namespace qttabbar::hooks

