#pragma once

#include <mutex>
#include <optional>
#include <vector>

#include "PluginLibrary.h"

namespace qttabbar::plugins {

class PluginManagerNative {
public:
    static PluginManagerNative& Instance();

    void Refresh();
    std::vector<PluginMetadataNative> EnumerateMetadata() const;
    bool SetEnabled(const std::wstring& pluginId, bool enabled);
    bool IsEnabled(const std::wstring& pluginId) const;

private:
    PluginManagerNative();

    void LoadFromRegistry();
    void LoadEnabledList();
    void PersistEnabledList() const;
    PluginLibrary* FindLibraryById(const std::wstring& pluginId);
    const PluginLibrary* FindLibraryById(const std::wstring& pluginId) const;

    std::wstring MakePluginId(const PluginMetadataNative& metadata) const;

    mutable std::mutex mutex_;
    std::vector<PluginLibrary> libraries_;
    std::vector<std::wstring> enabledIds_;
    bool initialized_ = false;
};

}  // namespace qttabbar::plugins

