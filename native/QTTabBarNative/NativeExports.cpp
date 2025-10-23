#include "pch.h"

#include "PluginManagerNative.h"
#include "PluginContracts.h"

#include "HookLibraryBridge.h"

using namespace qttabbar::plugins;

extern "C" {

__declspec(dllexport) void __stdcall QTTabBarNative_RefreshPlugins() {
    PluginManagerNative::Instance().Refresh();
}

__declspec(dllexport) size_t __stdcall QTTabBarNative_GetPluginCount() {
    return PluginManagerNative::Instance().EnumerateMetadata().size();
}

__declspec(dllexport) BOOL __stdcall QTTabBarNative_GetPluginMetadata(size_t index, PluginMetadataNative* metadata) {
    if (metadata == nullptr) {
        return FALSE;
    }
    auto list = PluginManagerNative::Instance().EnumerateMetadata();
    if (index >= list.size()) {
        return FALSE;
    }
    *metadata = list[index];
    return TRUE;
}

__declspec(dllexport) BOOL __stdcall QTTabBarNative_SetPluginEnabled(const wchar_t* pluginId, BOOL enabled) {
    if (pluginId == nullptr) {
        return FALSE;
    }
    return PluginManagerNative::Instance().SetEnabled(pluginId, enabled != FALSE) ? TRUE : FALSE;
}

__declspec(dllexport) BOOL __stdcall QTTabBarNative_IsPluginEnabled(const wchar_t* pluginId) {
    if (pluginId == nullptr) {
        return FALSE;
    }
    return PluginManagerNative::Instance().IsEnabled(pluginId) ? TRUE : FALSE;
}

__declspec(dllexport) int __stdcall QTTabBarNative_InitializeHookLibrary(const qttabbar::hooks::HookCallbacks* callbacks,
                                                                         const wchar_t* libraryPath) {
    if (callbacks == nullptr) {
        return E_POINTER;
    }
    qttabbar::hooks::HookCallbacks local = *callbacks;
    return qttabbar::hooks::HookLibraryBridge::Instance().Initialize(local, libraryPath);
}

__declspec(dllexport) void __stdcall QTTabBarNative_ShutdownHookLibrary() {
    qttabbar::hooks::HookLibraryBridge::Instance().Shutdown();
}

__declspec(dllexport) int __stdcall QTTabBarNative_InitShellBrowserHook(IUnknown* shellBrowser) {
    if (shellBrowser == nullptr) {
        return E_POINTER;
    }
    return qttabbar::hooks::HookLibraryBridge::Instance().InitShellBrowserHook(shellBrowser);
}

}  // extern "C"
