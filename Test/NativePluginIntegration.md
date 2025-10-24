# Native Plugin Integration Test Plan

These manual regression tests validate the new native plugin hosting APIs that mirror the legacy managed loader.

## Prerequisites
* Install the built `QTTabBarNative.dll` together with the managed host on a test machine.
* Ensure default plugins are deployed under `%ProgramData%\QTTabBar` (e.g. `QTQuick.dll`, `QTClock.dll`, etc.).
* Start an elevated PowerShell session for registry inspection as needed.

## Test Cases

1. **Plugin enumeration parity**
   1. Launch a small C# harness that calls `QTTabBarNative_RefreshPlugins` followed by `QTTabBarNative_GetPluginCount`/`QTTabBarNative_GetPluginMetadata`.
   2. Verify every DLL located in `%ProgramData%\QTTabBar` is reported with matching metadata (name, version, managed fallback flag).
   3. Confirm the order matches the persisted `HKCU\Software\QTTabBar\Plugins\Paths` list.

2. **Enable/disable persistence**
   1. Call `QTTabBarNative_SetPluginEnabled` for two plugins, toggling one on and one off.
   2. Confirm `QTTabBarNative_IsPluginEnabled` reflects the new state immediately.
   3. Inspect `HKCU\Software\QTTabBar\Plugins\Enabled` and ensure the list mirrors the toggles without removing unrelated entries.

3. **Native instantiation**
   1. Choose a plugin where `managedFallback == false`.
   2. Call `QTTabBarNative_CreatePluginInstance` and assert the returned handle and `PluginClientVTable` pointer are non-null.
   3. Invoke `QTTabBarNative_PluginOpen` with mock `IPluginServer`/`IShellBrowser` stubs (can be simple `IUnknown` mocks).
   4. Trigger `QTTabBarNative_PluginOnMenuClick`, `QTTabBarNative_PluginOnOption`, and `QTTabBarNative_PluginOnShortcut` and verify breakpoints/logging in the plugin fire.
   5. Call `QTTabBarNative_PluginQueryShortcuts` and ensure the returned strings propagate back to managed code.
   6. Finish with `QTTabBarNative_PluginClose` (using `PluginEndCode::Unloaded`) and `QTTabBarNative_DestroyPluginInstance`.

4. **Managed fallback handling**
   1. Pick a plugin with `managedFallback == true`.
   2. Call `QTTabBarNative_CreatePluginInstance` and confirm it returns `HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)` without altering the registry.
   3. Verify the managed loader still instantiates the plugin successfully.

5. **Registry path reconciliation**
   1. Remove a DLL from `%ProgramData%\QTTabBar` and call `QTTabBarNative_RefreshPlugins`.
   2. Confirm the missing DLL is dropped from `QTTabBarNative_GetPluginMetadata` output and the registry `Plugins\Paths` key.
   3. Restore the DLL, refresh again, and ensure it reappears in both metadata enumeration and the registry list without duplicating entries.

Record pass/fail results and any anomalies in the QA tracker. These tests guarantee that the managed and native loaders stay interoperable across upgrades.
