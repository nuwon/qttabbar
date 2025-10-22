#include "Config.h"

#include <Windows.h>
#include <ShlObj.h>
#include <algorithm>
#include <gdiplus.h>
#include <sstream>

#include "third_party/nlohmann/json.hpp"

#pragma comment(lib, "Gdiplus.lib")

namespace qttabbar {
namespace {

using nlohmann::json;

constexpr const wchar_t kRegRoot[] = L"Software\\QTTabBar\\Config";

struct GdiplusToken {
    ULONG_PTR token = 0;
    GdiplusToken() {
        Gdiplus::GdiplusStartupInput input;
        if (Gdiplus::GdiplusStartup(&token, &input, nullptr) != Gdiplus::Ok) {
            token = 0;
        }
    }
    ~GdiplusToken() {
        if (token != 0) {
            Gdiplus::GdiplusShutdown(token);
        }
    }
};

uint32_t MakeColor(uint32_t rgb) {
    return 0xFF000000 | (rgb & 0x00FFFFFF);
}

StringList DefaultTextExtensions() {
    return {
        L".txt", L".rtf", L".ini", L".inf", L".properties", L".ruleset", L".settings",
        L".cs", L".log", L".js", L".vbs", L".bat", L".cmd", L".sh", L".c", L".cpp",
        L".cc", L".h", L".rc", L".xml", L".yml", L".yaml", L".htm", L".html", L".mht",
        L".mhtml", L".shtml", L".hta", L".hxt", L".hxc", L".hhc", L".hhk", L".hhp",
        L".java", L".sql", L".csv", L".md", L".m", L".reg", L".wxl", L".wxs", L".py",
        L".rb", L".jsp", L".asp", L".php", L".aspx", L".resx", L".xaml", L".config",
        L".manifest", L".csproj", L".vbproj"
    };
}

StringList SplitSemicolonList(const std::wstring& value) {
    StringList result;
    size_t start = 0;
    while (start < value.size()) {
        size_t end = value.find(L';', start);
        if (end == std::wstring::npos) {
            end = value.size();
        }
        if (end > start) {
            std::wstring token = value.substr(start, end - start);
            if (!token.empty()) {
                std::transform(token.begin(), token.end(), token.begin(), ::towlower);
                result.push_back(token);
            }
        }
        start = end + 1;
    }
    return result;
}

StringList DefaultImageExtensions() {
    GdiplusToken token;
    UINT count = 0;
    UINT bytes = 0;
    StringList result;
    if (Gdiplus::GetImageDecodersSize(&count, &bytes) == Gdiplus::Ok && bytes != 0) {
        std::vector<uint8_t> buffer(bytes);
        auto* decoders = reinterpret_cast<Gdiplus::ImageCodecInfo*>(buffer.data());
        if (Gdiplus::GetImageDecoders(count, bytes, decoders) == Gdiplus::Ok) {
            for (UINT i = 0; i < count; ++i) {
                const WCHAR* ext = decoders[i].FilenameExtension;
                if (!ext) continue;
                StringList parts = SplitSemicolonList(ext);
                result.insert(result.end(), parts.begin(), parts.end());
            }
        }
    }
    static const wchar_t kSupportedMovies[] =
        L".asx;.dvr-ms;.mp2;.flv;.mkv;.ts;.3g2;.3gp;.3gp2;.3gpp;.amr;.amv;.asf;.avi;.bdmv;.bik;"
        L".d2v;.divx;.drc;.dsa;.dsm;.dss;.dsv;.evo;.f4v;.flc;.fli;.flic;.flv;.hdmov;.ifo;.ivf;.m1v;.m2p;.m2t;"
        L".m2ts;.m2v;.m4b;.m4p;.m4v;.mkv;.mp2v;.mp4;.mp4v;.mpe;.mpeg;.mpg;.mpls;.mpv2;.mpv4;.mov;.mts;.ogm;.ogv;"
        L".pss;.pva;.qt;.ram;.ratdvd;.rm;.rmm;.rmvb;.roq;.rpm;.smil;.smk;.swf;.tp;.tpr;.ts;.vob;.vp6;.webm;.wm;.wmp;.wmv";
    StringList movieExt = SplitSemicolonList(kSupportedMovies);
    result.insert(result.end(), movieExt.begin(), movieExt.end());
    std::sort(result.begin(), result.end());
    result.erase(std::unique(result.begin(), result.end()), result.end());
    return result;
}

ByteVector PidlFromDisplayName(const wchar_t* name) {
    ByteVector result;
    PIDLIST_RELATIVE pidl = nullptr;
    ULONG eaten = 0;
    HRESULT hr = SHParseDisplayName(name, nullptr, &pidl, 0, &eaten);
    if (SUCCEEDED(hr) && pidl) {
        size_t size = ILGetSize(pidl);
        result.resize(size);
        std::copy_n(reinterpret_cast<const uint8_t*>(pidl), size, result.begin());
    }
    if (pidl) {
        CoTaskMemFree(pidl);
    }
    return result;
}

bool ReadDwordValue(HKEY key, const wchar_t* name, DWORD* out) {
    DWORD data = 0;
    DWORD size = sizeof(DWORD);
    DWORD type = 0;
    if (RegGetValueW(key, nullptr, name, RRF_RT_REG_DWORD, &type, &data, &size) == ERROR_SUCCESS) {
        *out = data;
        return true;
    }
    return false;
}

bool ReadStringValue(HKEY key, const wchar_t* name, std::wstring* out) {
    DWORD type = 0;
    DWORD size = 0;
    LONG status = RegGetValueW(key, nullptr, name, RRF_RT_REG_SZ, &type, nullptr, &size);
    if (status != ERROR_SUCCESS) {
        return false;
    }
    std::wstring buffer;
    buffer.resize(size / sizeof(wchar_t));
    status = RegGetValueW(key, nullptr, name, RRF_RT_REG_SZ, &type, buffer.data(), &size);
    if (status != ERROR_SUCCESS) {
        return false;
    }
    if (!buffer.empty() && buffer.back() == L'\0') {
        buffer.pop_back();
    }
    *out = std::move(buffer);
    return true;
}

void WriteDwordValue(HKEY key, const wchar_t* name, DWORD value) {
    RegSetValueExW(key, name, 0, REG_DWORD, reinterpret_cast<const BYTE*>(&value), sizeof(value));
}

void WriteStringValue(HKEY key, const wchar_t* name, const std::wstring& value) {
    const BYTE* data = reinterpret_cast<const BYTE*>(value.c_str());
    DWORD size = static_cast<DWORD>((value.size() + 1) * sizeof(wchar_t));
    RegSetValueExW(key, name, 0, REG_SZ, data, size);
}

void WriteJsonValue(HKEY key, const wchar_t* name, const json& j) {
    std::string utf8 = j.dump();
    int length = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), static_cast<int>(utf8.size()), nullptr, 0);
    std::wstring wide(length, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), static_cast<int>(utf8.size()), wide.data(), length);
    WriteStringValue(key, name, wide);
}

std::optional<json> ReadJsonValue(HKEY key, const wchar_t* name) {
    std::wstring str;
    if (!ReadStringValue(key, name, &str)) {
        return std::nullopt;
    }
    int length = WideCharToMultiByte(CP_UTF8, 0, str.c_str(), static_cast<int>(str.size()), nullptr, 0, nullptr, nullptr);
    std::string utf8(length, '\0');
    WideCharToMultiByte(CP_UTF8, 0, str.c_str(), static_cast<int>(str.size()), utf8.data(), length, nullptr, nullptr);
    try {
        return json::parse(utf8);
    } catch (...) {
        return std::nullopt;
    }
}

ColorValue ParseColor(const json& j) {
    if (j.is_object()) {
        if (j.contains("Argb")) {
            return ColorValue(static_cast<uint32_t>(j["Argb"].get<int64_t>()));
        }
        if (j.contains("Value")) {
            return ColorValue(static_cast<uint32_t>(j["Value"].get<int64_t>()));
        }
        if (j.contains("A") && j.contains("R") && j.contains("G") && j.contains("B")) {
            uint32_t value = (j["A"].get<uint32_t>() << 24) |
                              (j["R"].get<uint32_t>() << 16) |
                              (j["G"].get<uint32_t>() << 8) |
                              (j["B"].get<uint32_t>());
            return ColorValue(value);
        }
    } else if (j.is_number_unsigned()) {
        return ColorValue(static_cast<uint32_t>(j.get<uint64_t>()));
    }
    return ColorValue();
}

json SerializeColor(ColorValue color) {
    json j;
    j["Value"] = color.argb;
    return j;
}

Padding ParsePadding(const json& j) {
    Padding padding;
    if (j.is_object()) {
        padding.left = j.value("Left", 0);
        padding.top = j.value("Top", 0);
        padding.right = j.value("Right", 0);
        padding.bottom = j.value("Bottom", 0);
    } else if (j.is_array() && j.size() == 4) {
        padding.left = j[0].get<int>();
        padding.top = j[1].get<int>();
        padding.right = j[2].get<int>();
        padding.bottom = j[3].get<int>();
    }
    return padding;
}

json SerializePadding(const Padding& padding) {
    json j;
    j["Left"] = padding.left;
    j["Top"] = padding.top;
    j["Right"] = padding.right;
    j["Bottom"] = padding.bottom;
    return j;
}

FontConfig ParseFont(const json& j) {
    FontConfig font;
    if (j.is_object()) {
        if (j.contains("FontName")) {
            font.family = std::wstring(j["FontName"].get<std::string>().begin(), j["FontName"].get<std::string>().end());
        } else if (j.contains("Name")) {
            font.family = std::wstring(j["Name"].get<std::string>().begin(), j["Name"].get<std::string>().end());
        }
        font.size = j.value("FontSize", j.value("Size", 9.0f));
        font.style = j.value("FontStyle", j.value("Style", 0));
    }
    return font;
}

json SerializeFont(const FontConfig& font) {
    json j;
    std::wstring name = font.family.empty() ? L"微软雅黑" : font.family;
    std::string utf8Name;
    int length = WideCharToMultiByte(CP_UTF8, 0, name.c_str(), static_cast<int>(name.size()), nullptr, 0, nullptr, nullptr);
    utf8Name.resize(length);
    WideCharToMultiByte(CP_UTF8, 0, name.c_str(), static_cast<int>(name.size()), utf8Name.data(), length, nullptr, nullptr);
    j["FontName"] = utf8Name;
    j["FontSize"] = font.size;
    j["FontStyle"] = font.style;
    return j;
}

StringList ParseStringList(const json& j) {
    StringList list;
    if (j.is_array()) {
        for (const auto& item : j) {
            if (item.is_string()) {
                std::string utf8 = item.get<std::string>();
                int len = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), static_cast<int>(utf8.size()), nullptr, 0);
                std::wstring wide(len, L'\0');
                MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), static_cast<int>(utf8.size()), wide.data(), len);
                list.push_back(std::move(wide));
            }
        }
    }
    return list;
}

json SerializeStringList(const StringList& list) {
    json j = json::array();
    for (const auto& item : list) {
        int len = WideCharToMultiByte(CP_UTF8, 0, item.c_str(), static_cast<int>(item.size()), nullptr, 0, nullptr, nullptr);
        std::string utf8(len, '\0');
        WideCharToMultiByte(CP_UTF8, 0, item.c_str(), static_cast<int>(item.size()), utf8.data(), len, nullptr, nullptr);
        j.push_back(utf8);
    }
    return j;
}

MouseActionMap ParseMouseActionMap(const json& j) {
    MouseActionMap map;
    if (j.is_object()) {
        for (auto it = j.begin(); it != j.end(); ++it) {
            MouseChord key = static_cast<MouseChord>(std::stoul(it.key()));
            BindAction value = static_cast<BindAction>(it.value().get<uint32_t>());
            map[key] = value;
        }
    } else if (j.is_array()) {
        for (const auto& entry : j) {
            MouseChord key = static_cast<MouseChord>(entry.value("Key", 0));
            BindAction value = static_cast<BindAction>(entry.value("Value", 0));
            map[key] = value;
        }
    }
    return map;
}

json SerializeMouseActionMap(const MouseActionMap& map) {
    json j = json::array();
    for (const auto& pair : map) {
        json entry;
        entry["Key"] = static_cast<uint32_t>(pair.first);
        entry["Value"] = static_cast<uint32_t>(pair.second);
        j.push_back(entry);
    }
    return j;
}

IntVector ParseIntVector(const json& j) {
    IntVector vec;
    if (j.is_array()) {
        for (const auto& entry : j) {
            vec.push_back(entry.get<int>());
        }
    }
    return vec;
}


ByteVector ParseByteVector(const json& j) {
    ByteVector data;
    if (j.is_array()) {
        for (const auto& value : j) {
            int v = value.get<int>();
            data.push_back(static_cast<uint8_t>(std::clamp(v, 0, 255)));
        }
    }
    return data;
}

json SerializeByteVector(const ByteVector& data) {
    json j = json::array();
    for (uint8_t v : data) {
        j.push_back(static_cast<int>(v));
    }
    return j;
}
json SerializeIntVector(const IntVector& vec) {
    json j = json::array();
    for (int value : vec) {
        j.push_back(value);
    }
    return j;
}

std::map<std::wstring, IntVector> ParsePluginShortcuts(const json& j) {
    std::map<std::wstring, IntVector> map;
    if (j.is_object()) {
        for (auto it = j.begin(); it != j.end(); ++it) {
            std::string keyUtf8 = it.key();
            int len = MultiByteToWideChar(CP_UTF8, 0, keyUtf8.c_str(), static_cast<int>(keyUtf8.size()), nullptr, 0);
            std::wstring key(len, L'\0');
            MultiByteToWideChar(CP_UTF8, 0, keyUtf8.c_str(), static_cast<int>(keyUtf8.size()), key.data(), len);
            map.emplace(std::move(key), ParseIntVector(it.value()));
        }
    }
    return map;
}

json SerializePluginShortcuts(const std::map<std::wstring, IntVector>& map) {
    json j;
    for (const auto& pair : map) {
        int len = WideCharToMultiByte(CP_UTF8, 0, pair.first.c_str(), static_cast<int>(pair.first.size()), nullptr, 0, nullptr, nullptr);
        std::string key(len, '\0');
        WideCharToMultiByte(CP_UTF8, 0, pair.first.c_str(), static_cast<int>(pair.first.size()), key.data(), len, nullptr, nullptr);
        j[key] = SerializeIntVector(pair.second);
    }
    return j;
}

}  // namespace

ConfigData::ConfigData() {
    window.captureWeChatSelection = true;
    window.defaultLocation = PidlFromDisplayName(L"::{20D04FE0-3AEA-1069-A2D8-08002B30309D}");
    tweaks.altRowBackgroundColor = ColorValue(MakeColor(0xfaf5f1));
    tips.textExt = DefaultTextExtensions();
    tips.imageExt = DefaultImageExtensions();

    // Button indexes default mirrors managed behavior.
    if (LOBYTE(LOWORD(GetVersion())) < 6) {
        bbar.buttonIndexes = {1, 2, 3, 4, 5, 6, 7, 11, 13, 12, 14, 15, 21, 9, 20};
    } else {
        bbar.buttonIndexes = {3, 4, 5, 6, 7, 17, 11, 12, 14, 15, 13, 21, 9, 19, 10};
    }

    mouse.globalMouseActions = {
        {MouseChord::X1, BindAction::GoBack},
        {MouseChord::X2, BindAction::GoForward},
        {MouseChord::X1 | MouseChord::Ctrl, BindAction::GoFirst},
        {MouseChord::X2 | MouseChord::Ctrl, BindAction::GoLast},
    };
    mouse.tabActions = {
        {MouseChord::Middle, BindAction::CloseTab},
        {MouseChord::Ctrl | MouseChord::Left, BindAction::LockTab},
        {MouseChord::Double, BindAction::UpOneLevelTab},
    };
    mouse.barActions = {
        {MouseChord::Double, BindAction::NewTab},
        {MouseChord::Middle, BindAction::RestoreLastClosed},
        {MouseChord::Ctrl | MouseChord::Middle, BindAction::TearOffCurrent},
    };
    mouse.linkActions = {
        {MouseChord::None, BindAction::ItemsOpenInNewTabNoSel},
        {MouseChord::Middle, BindAction::ItemOpenInNewTab},
        {MouseChord::Ctrl | MouseChord::Middle, BindAction::ItemOpenInNewWindow},
    };
    mouse.itemActions = {
        {MouseChord::Middle, BindAction::ItemOpenInNewTab},
        {MouseChord::Ctrl | MouseChord::Middle, BindAction::ItemsOpenInNewTabNoSel},
    };
    mouse.marginActions = {
        {MouseChord::Double, BindAction::UpOneLevel},
        {MouseChord::Middle, BindAction::BrowseFolder},
        {MouseChord::Ctrl | MouseChord::Double, BindAction::OpenCmd},
        {MouseChord::Ctrl | MouseChord::Middle, BindAction::ItemsOpenInNewTabNoSel},
    };

    keys.shortcuts.resize(static_cast<size_t>(BindAction::KEYBOARD_ACTION_COUNT));
    auto setShortcut = [this](BindAction action, int value) {
        size_t index = static_cast<size_t>(action);
        if (index >= keys.shortcuts.size()) {
            keys.shortcuts.resize(index + 1);
        }
        keys.shortcuts[index] = value;
    };
    const int enabledFlag = 1 << 30;
    auto encode = [enabledFlag](int key) { return key | enabledFlag; };
    setShortcut(BindAction::GoBack, encode(VK_LEFT | (MOD_ALT << 16)));
    setShortcut(BindAction::GoForward, encode(VK_RIGHT | (MOD_ALT << 16)));
    setShortcut(BindAction::GoFirst, encode(VK_LEFT | (MOD_ALT << 16) | (MOD_CONTROL << 16)));
    setShortcut(BindAction::GoLast, encode(VK_RIGHT | (MOD_ALT << 16) | (MOD_CONTROL << 16)));
    setShortcut(BindAction::NextTab, encode(VK_TAB | (MOD_CONTROL << 16)));
    setShortcut(BindAction::PreviousTab, encode(VK_TAB | (MOD_CONTROL << 16) | (MOD_SHIFT << 16)));
    setShortcut(BindAction::NewTab, encode('T' | (MOD_CONTROL << 16)));
    setShortcut(BindAction::NewWindow, encode('T' | (MOD_CONTROL << 16) | (MOD_SHIFT << 16)));
    setShortcut(BindAction::CloseCurrent, encode('W' | (MOD_CONTROL << 16)));
    setShortcut(BindAction::CloseAllButCurrent, encode('W' | (MOD_CONTROL << 16) | (MOD_SHIFT << 16)));
    setShortcut(BindAction::RestoreLastClosed, encode('Z' | (MOD_CONTROL << 16) | (MOD_SHIFT << 16)));
    setShortcut(BindAction::BrowseFolder, encode('O' | (MOD_CONTROL << 16)));
    setShortcut(BindAction::ShowOptions, encode('O' | (MOD_ALT << 16)));
    setShortcut(BindAction::ShowToolbarMenu, encode(VK_OEM_COMMA | (MOD_ALT << 16)));
    setShortcut(BindAction::ShowTabMenuCurrent, encode(VK_OEM_PERIOD | (MOD_ALT << 16)));
    setShortcut(BindAction::ShowGroupMenu, encode('G' | (MOD_ALT << 16)));
    setShortcut(BindAction::ShowUserAppsMenu, encode('H' | (MOD_ALT << 16)));
    setShortcut(BindAction::ShowRecentTabsMenu, encode('U' | (MOD_ALT << 16)));
    setShortcut(BindAction::ShowRecentFilesMenu, encode('F' | (MOD_ALT << 16)));
    setShortcut(BindAction::NewFile, encode('N' | (MOD_CONTROL << 16) | (MOD_ALT << 16)));
    setShortcut(BindAction::CreateNewGroup, encode('D' | (MOD_CONTROL << 16)));

    plugin.enabled.clear();
    lang.pluginLangFiles.clear();
}

static std::wstring MakeCategoryPath(const wchar_t* category) {
    std::wstring path = kRegRoot;
    path.append(L"\\");
    path.append(category);
    return path;
}

void ReadWindowSettings(HKEY key, WindowSettings& settings) {
    DWORD value = 0;
    if (ReadDwordValue(key, L"CaptureNewWindows", &value)) settings.captureNewWindows = value != 0;
    if (ReadDwordValue(key, L"CaptureWeChatSelection", &value)) settings.captureWeChatSelection = value != 0;
    if (ReadDwordValue(key, L"RestoreSession", &value)) settings.restoreSession = value != 0;
    if (ReadDwordValue(key, L"RestoreOnlyLocked", &value)) settings.restoreOnlyLocked = value != 0;
    if (ReadDwordValue(key, L"CloseBtnClosesUnlocked", &value)) settings.closeBtnClosesUnlocked = value != 0;
    if (ReadDwordValue(key, L"CloseBtnClosesSingleTab", &value)) settings.closeBtnClosesSingleTab = value != 0;
    if (ReadDwordValue(key, L"TrayOnClose", &value)) settings.trayOnClose = value != 0;
    if (ReadDwordValue(key, L"TrayOnMinimize", &value)) settings.trayOnMinimize = value != 0;
    if (ReadDwordValue(key, L"AutoHookWindow", &value)) settings.autoHookWindow = value != 0;
    if (ReadDwordValue(key, L"ShowFailNavMsg", &value)) settings.showFailNavMsg = value != 0;
    if (auto jsonValue = ReadJsonValue(key, L"DefaultLocation")) {
        settings.defaultLocation = ParseByteVector(*jsonValue);
    }
}

void WriteWindowSettings(HKEY key, const WindowSettings& settings) {
    WriteDwordValue(key, L"CaptureNewWindows", (settings.captureNewWindows ? 1u : 0u));
    WriteDwordValue(key, L"CaptureWeChatSelection", (settings.captureWeChatSelection ? 1u : 0u));
    WriteDwordValue(key, L"RestoreSession", (settings.restoreSession ? 1u : 0u));
    WriteDwordValue(key, L"RestoreOnlyLocked", (settings.restoreOnlyLocked ? 1u : 0u));
    WriteDwordValue(key, L"CloseBtnClosesUnlocked", (settings.closeBtnClosesUnlocked ? 1u : 0u));
    WriteDwordValue(key, L"CloseBtnClosesSingleTab", (settings.closeBtnClosesSingleTab ? 1u : 0u));
    WriteDwordValue(key, L"TrayOnClose", (settings.trayOnClose ? 1u : 0u));
    WriteDwordValue(key, L"TrayOnMinimize", (settings.trayOnMinimize ? 1u : 0u));
    WriteDwordValue(key, L"AutoHookWindow", (settings.autoHookWindow ? 1u : 0u));
    WriteDwordValue(key, L"ShowFailNavMsg", (settings.showFailNavMsg ? 1u : 0u));
    json value = SerializeByteVector(settings.defaultLocation);
    WriteJsonValue(key, L"DefaultLocation", value);
}

void ReadTabsSettings(HKEY key, TabsSettings& settings) {
    DWORD value = 0;
    if (ReadDwordValue(key, L"NewTabPosition", &value)) settings.newTabPosition = static_cast<TabPos>(value);
    if (ReadDwordValue(key, L"NextAfterClosed", &value)) settings.nextAfterClosed = static_cast<TabPos>(value);
    if (ReadDwordValue(key, L"ActivateNewTab", &value)) settings.activateNewTab = value != 0;
    if (ReadDwordValue(key, L"NeverOpenSame", &value)) settings.neverOpenSame = value != 0;
    if (ReadDwordValue(key, L"RenameAmbTabs", &value)) settings.renameAmbTabs = value != 0;
    if (ReadDwordValue(key, L"DragOverTabOpensSDT", &value)) settings.dragOverTabOpensSDT = value != 0;
    if (ReadDwordValue(key, L"ShowFolderIcon", &value)) settings.showFolderIcon = value != 0;
    if (ReadDwordValue(key, L"ShowSubDirTipOnTab", &value)) settings.showSubDirTipOnTab = value != 0;
    if (ReadDwordValue(key, L"ShowDriveLetters", &value)) settings.showDriveLetters = value != 0;
    if (ReadDwordValue(key, L"ShowCloseButtons", &value)) settings.showCloseButtons = value != 0;
    if (ReadDwordValue(key, L"CloseBtnsWithAlt", &value)) settings.closeBtnsWithAlt = value != 0;
    if (ReadDwordValue(key, L"CloseBtnsOnHover", &value)) settings.closeBtnsOnHover = value != 0;
    if (ReadDwordValue(key, L"ShowNavButtons", &value)) settings.showNavButtons = value != 0;
    if (ReadDwordValue(key, L"NavButtonsOnRight", &value)) settings.navButtonsOnRight = value != 0;
    if (ReadDwordValue(key, L"MultipleTabRows", &value)) settings.multipleTabRows = value != 0;
    if (ReadDwordValue(key, L"ActiveTabOnBottomRow", &value)) settings.activeTabOnBottomRow = value != 0;
    if (ReadDwordValue(key, L"NeedPlusButton", &value)) settings.needPlusButton = value != 0;
}

void WriteTabsSettings(HKEY key, const TabsSettings& settings) {
    WriteDwordValue(key, L"NewTabPosition", static_cast<DWORD>(settings.newTabPosition));
    WriteDwordValue(key, L"NextAfterClosed", static_cast<DWORD>(settings.nextAfterClosed));
    WriteDwordValue(key, L"ActivateNewTab", (settings.activateNewTab ? 1u : 0u));
    WriteDwordValue(key, L"NeverOpenSame", (settings.neverOpenSame ? 1u : 0u));
    WriteDwordValue(key, L"RenameAmbTabs", (settings.renameAmbTabs ? 1u : 0u));
    WriteDwordValue(key, L"DragOverTabOpensSDT", (settings.dragOverTabOpensSDT ? 1u : 0u));
    WriteDwordValue(key, L"ShowFolderIcon", (settings.showFolderIcon ? 1u : 0u));
    WriteDwordValue(key, L"ShowSubDirTipOnTab", (settings.showSubDirTipOnTab ? 1u : 0u));
    WriteDwordValue(key, L"ShowDriveLetters", (settings.showDriveLetters ? 1u : 0u));
    WriteDwordValue(key, L"ShowCloseButtons", (settings.showCloseButtons ? 1u : 0u));
    WriteDwordValue(key, L"CloseBtnsWithAlt", (settings.closeBtnsWithAlt ? 1u : 0u));
    WriteDwordValue(key, L"CloseBtnsOnHover", (settings.closeBtnsOnHover ? 1u : 0u));
    WriteDwordValue(key, L"ShowNavButtons", (settings.showNavButtons ? 1u : 0u));
    WriteDwordValue(key, L"NavButtonsOnRight", (settings.navButtonsOnRight ? 1u : 0u));
    WriteDwordValue(key, L"MultipleTabRows", (settings.multipleTabRows ? 1u : 0u));
    WriteDwordValue(key, L"ActiveTabOnBottomRow", (settings.activeTabOnBottomRow ? 1u : 0u));
    WriteDwordValue(key, L"NeedPlusButton", (settings.needPlusButton ? 1u : 0u));
}

}  // namespace

ConfigData LoadConfigFromRegistry() {
    ConfigData config;
    HKEY root = nullptr;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, kRegRoot, 0, KEY_READ, &root) != ERROR_SUCCESS) {
        return config;
    }

    auto readCategory = [&](const wchar_t* name, auto&& reader, auto& settings) {
        std::wstring path = MakeCategoryPath(name);
        HKEY key = nullptr;
        if (RegOpenKeyExW(HKEY_CURRENT_USER, path.c_str(), 0, KEY_READ, &key) == ERROR_SUCCESS) {
            reader(key, settings);
            RegCloseKey(key);
        }
    };

    readCategory(L"Window", ReadWindowSettings, config.window);
    readCategory(L"Tabs", ReadTabsSettings, config.tabs);

    RegCloseKey(root);
    return config;
}

void WriteConfigToRegistry(const ConfigData& config, bool desktopOnly) {
    HKEY root = nullptr;
    RegCreateKeyExW(HKEY_CURRENT_USER, kRegRoot, 0, nullptr, 0, KEY_WRITE, nullptr, &root, nullptr);
    auto writeCategory = [&](const wchar_t* name, auto&& writer, const auto& settings) {
        std::wstring path = MakeCategoryPath(name);
        HKEY key = nullptr;
        if (RegCreateKeyExW(HKEY_CURRENT_USER, path.c_str(), 0, nullptr, 0, KEY_WRITE, nullptr, &key, nullptr) == ERROR_SUCCESS) {
            writer(key, settings);
            RegCloseKey(key);
        }
    };

    if (!desktopOnly) {
        writeCategory(L"Window", WriteWindowSettings, config.window);
        writeCategory(L"Tabs", WriteTabsSettings, config.tabs);
    } else {
        writeCategory(L"Desktop", [](HKEY, const DesktopSettings&) {}, config.desktop);
    }

    if (root) {
        RegCloseKey(root);
    }
}

void UpdateConfigSideEffects(const ConfigData&, bool broadcastChanges) {
    if(!broadcastChanges) {
        return;
    }
    static constexpr wchar_t kNotificationName[] = L"QTTabBar.ConfigChanged";
    ::SendMessageTimeoutW(HWND_BROADCAST, WM_SETTINGCHANGE, 0,
                          reinterpret_cast<LPARAM>(kNotificationName),
                          SMTO_ABORTIFHUNG, 100, nullptr);
}

}  // namespace qttabbar

