#include "pch.h"

#include "PluginLibrary.h"

#include <shlwapi.h>

#include <cstring>
#include <cwchar>
#include <memory>

#pragma comment(lib, "Shlwapi.lib")

namespace qttabbar::plugins {

namespace {
constexpr wchar_t kQueryMetadataProc[] = L"QTPlugin_QueryMetadata";
constexpr wchar_t kCreateProc[] = L"QTPlugin_CreateInstance";
constexpr wchar_t kDestroyProc[] = L"QTPlugin_DestroyInstance";
}  // namespace

PluginLibrary::PluginLibrary(std::wstring path, bool enabled)
    : path_(std::move(path)) {
    ResetMetadata();
    metadata_.enabled = enabled ? TRUE : FALSE;
}

PluginLibrary::PluginLibrary(PluginLibrary&& other) noexcept {
    *this = std::move(other);
}

PluginLibrary& PluginLibrary::operator=(PluginLibrary&& other) noexcept {
    if (this != &other) {
        Unload();
        path_ = std::move(other.path_);
        module_ = other.module_;
        other.module_ = nullptr;
        metadata_ = other.metadata_;
        exports_ = other.exports_;
        other.ResetMetadata();
    }
    return *this;
}

PluginLibrary::~PluginLibrary() {
    Unload();
}

void PluginLibrary::ResetMetadata() {
    std::memset(&metadata_, 0, sizeof(metadata_));
    metadata_.type = PluginType::Interactive;
}

bool PluginLibrary::Load() {
    if (module_ != nullptr) {
        return true;
    }
    ResetMetadata();
    exports_ = {};

    if (path_.empty() || !PathFileExistsW(path_.c_str())) {
        return false;
    }

    module_ = ::LoadLibraryExW(path_.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DEFAULT_DIRS);
    if (module_ == nullptr) {
        return false;
    }

    exports_.queryMetadata = reinterpret_cast<PluginQueryMetadataFn>(
        ::GetProcAddress(module_, kQueryMetadataProc));
    exports_.createInstance = reinterpret_cast<PluginCreateFn>(
        ::GetProcAddress(module_, kCreateProc));
    exports_.destroyInstance = reinterpret_cast<PluginDestroyFn>(
        ::GetProcAddress(module_, kDestroyProc));

    if (exports_.queryMetadata == nullptr) {
        Unload();
        return false;
    }

    if (FAILED(exports_.queryMetadata(&metadata_))) {
        Unload();
        return false;
    }

    if (metadata_.libraryPath[0] == L'\0') {
        wcsncpy_s(metadata_.libraryPath, path_.c_str(), _TRUNCATE);
    }

    return true;
}

void PluginLibrary::Unload() {
    if (module_ != nullptr) {
        ::FreeLibrary(module_);
        module_ = nullptr;
    }
    exports_ = {};
    ResetMetadata();
}

std::optional<PluginLibraryExports> PluginLibrary::Exports() const {
    if (module_ == nullptr) {
        return std::nullopt;
    }
    return exports_;
}

}  // namespace qttabbar::plugins

