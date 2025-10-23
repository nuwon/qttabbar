#include "pch.h"

#include "PluginManagerNative.h"

#include <shlwapi.h>

#include <algorithm>
#include <cwchar>
#include <functional>
#include <set>
#include <string_view>

#include "PluginContracts.h"

#pragma comment(lib, "Shlwapi.lib")

namespace qttabbar::plugins {

namespace {
constexpr wchar_t kRegistryRoot[] = L"Software\\QTTabBar";
constexpr wchar_t kPluginPathsKey[] = L"Plugins\\Paths";
constexpr wchar_t kPluginEnabledKey[] = L"Plugins\\Enabled";

std::wstring ReadRegistryString(HKEY key, const wchar_t* name) {
    DWORD type = 0;
    DWORD size = 0;
    if (RegQueryValueExW(key, name, nullptr, &type, nullptr, &size) != ERROR_SUCCESS) {
        return {};
    }
    if (type != REG_SZ && type != REG_EXPAND_SZ) {
        return {};
    }
    std::wstring value(size / sizeof(wchar_t), L'\0');
    if (RegQueryValueExW(key, name, nullptr, &type, reinterpret_cast<LPBYTE>(value.data()), &size) != ERROR_SUCCESS) {
        return {};
    }
    if (!value.empty() && value.back() == L'\0') {
        value.pop_back();
    }
    return value;
}

void WriteRegistryString(HKEY key, const wchar_t* name, const std::wstring& value) {
    DWORD size = static_cast<DWORD>((value.size() + 1) * sizeof(wchar_t));
    RegSetValueExW(key, name, 0, REG_SZ, reinterpret_cast<const BYTE*>(value.c_str()), size);
}

std::vector<std::wstring> ReadRegistryStrings(HKEY key) {
    DWORD index = 0;
    std::vector<std::wstring> values;
    wchar_t name[260];
    DWORD nameLength = ARRAYSIZE(name);
    while (RegEnumValueW(key, index++, name, &nameLength, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS) {
        values.emplace_back(ReadRegistryString(key, name));
        nameLength = ARRAYSIZE(name);
    }
    return values;
}

void WriteRegistryStrings(HKEY key, const std::vector<std::wstring>& values) {
    DWORD index = 0;
    for (const auto& value : values) {
        wchar_t name[32];
        _snwprintf_s(name, _TRUNCATE, L"%u", index++);
        WriteRegistryString(key, name, value);
    }
}

}  // namespace

PluginManagerNative& PluginManagerNative::Instance() {
    static PluginManagerNative instance;
    return instance;
}

PluginManagerNative::PluginManagerNative() = default;

void PluginManagerNative::Refresh() {
    std::scoped_lock lock(mutex_);
    LoadFromRegistry();
    LoadEnabledList();
    initialized_ = true;
}

std::vector<PluginMetadataNative> PluginManagerNative::EnumerateMetadata() const {
    std::scoped_lock lock(mutex_);
    std::vector<PluginMetadataNative> metadata;
    metadata.reserve(libraries_.size());
    for (const auto& lib : libraries_) {
        metadata.push_back(lib.Metadata());
    }
    return metadata;
}

bool PluginManagerNative::SetEnabled(const std::wstring& pluginId, bool enabled) {
    std::scoped_lock lock(mutex_);
    PluginLibrary* library = FindLibraryById(pluginId);
    if (library == nullptr) {
        return false;
    }
    library->SetEnabled(enabled);
    auto it = std::find(enabledIds_.begin(), enabledIds_.end(), pluginId);
    if (enabled) {
        if (it == enabledIds_.end()) {
            enabledIds_.push_back(pluginId);
        }
    } else {
        if (it != enabledIds_.end()) {
            enabledIds_.erase(it);
        }
    }
    PersistEnabledList();
    return true;
}

bool PluginManagerNative::IsEnabled(const std::wstring& pluginId) const {
    std::scoped_lock lock(mutex_);
    return std::any_of(enabledIds_.begin(), enabledIds_.end(), [&](const std::wstring& id) {
        return _wcsicmp(id.c_str(), pluginId.c_str()) == 0;
    });
}

void PluginManagerNative::LoadFromRegistry() {
    libraries_.clear();

    HKEY root = nullptr;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, kRegistryRoot, 0, KEY_READ, &root) != ERROR_SUCCESS) {
        return;
    }

    HKEY pathsKey = nullptr;
    if (RegOpenKeyExW(root, kPluginPathsKey, 0, KEY_READ, &pathsKey) != ERROR_SUCCESS) {
        RegCloseKey(root);
        return;
    }

    std::vector<std::wstring> paths = ReadRegistryStrings(pathsKey);
    RegCloseKey(pathsKey);
    RegCloseKey(root);

    for (const auto& path : paths) {
        PluginLibrary library(path, false);
        if (library.Load()) {
            libraries_.push_back(std::move(library));
        }
    }
}

void PluginManagerNative::LoadEnabledList() {
    enabledIds_.clear();

    HKEY root = nullptr;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, kRegistryRoot, 0, KEY_READ, &root) != ERROR_SUCCESS) {
        return;
    }

    HKEY enabledKey = nullptr;
    if (RegOpenKeyExW(root, kPluginEnabledKey, 0, KEY_READ, &enabledKey) != ERROR_SUCCESS) {
        RegCloseKey(root);
        return;
    }

    enabledIds_ = ReadRegistryStrings(enabledKey);
    RegCloseKey(enabledKey);
    RegCloseKey(root);

    for (auto& library : libraries_) {
        std::wstring id = MakePluginId(library.Metadata());
        library.SetEnabled(IsEnabled(id));
    }
}

void PluginManagerNative::PersistEnabledList() const {
    HKEY root = nullptr;
    RegCreateKeyExW(HKEY_CURRENT_USER, kRegistryRoot, 0, nullptr, 0, KEY_WRITE, nullptr, &root, nullptr);
    if (root == nullptr) {
        return;
    }
    HKEY enabledKey = nullptr;
    RegCreateKeyExW(root, kPluginEnabledKey, 0, nullptr, 0, KEY_WRITE, nullptr, &enabledKey, nullptr);
    if (enabledKey != nullptr) {
        // Clear existing values.
        DWORD index = 0;
        wchar_t name[260];
        DWORD nameLength = ARRAYSIZE(name);
        while (RegEnumValueW(enabledKey, index, name, &nameLength, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS) {
            RegDeleteValueW(enabledKey, name);
            nameLength = ARRAYSIZE(name);
        }
        WriteRegistryStrings(enabledKey, enabledIds_);
        RegCloseKey(enabledKey);
    }
    RegCloseKey(root);
}

PluginLibrary* PluginManagerNative::FindLibraryById(const std::wstring& pluginId) {
    auto it = std::find_if(libraries_.begin(), libraries_.end(), [&](PluginLibrary& library) {
        return _wcsicmp(MakePluginId(library.Metadata()).c_str(), pluginId.c_str()) == 0;
    });
    return it != libraries_.end() ? &(*it) : nullptr;
}

const PluginLibrary* PluginManagerNative::FindLibraryById(const std::wstring& pluginId) const {
    auto it = std::find_if(libraries_.begin(), libraries_.end(), [&](const PluginLibrary& library) {
        return _wcsicmp(MakePluginId(library.Metadata()).c_str(), pluginId.c_str()) == 0;
    });
    return it != libraries_.end() ? &(*it) : nullptr;
}

std::wstring PluginManagerNative::MakePluginId(const PluginMetadataNative& metadata) const {
    std::wstring_view pathView(metadata.libraryPath);
    size_t hashValue = std::hash<std::wstring_view>{}(pathView);
    wchar_t hashBuffer[32] = {};
    _snwprintf_s(hashBuffer, _TRUNCATE, L"%IX", static_cast<unsigned int>(hashValue));

    std::wstring id = metadata.name;
    id.append(metadata.version);
    id.push_back(L'(');
    id.append(hashBuffer);
    id.push_back(L')');
    return id;
}

}  // namespace qttabbar::plugins

