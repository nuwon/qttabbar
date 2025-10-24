#include "pch.h"
#include "TabBarHost.h"

#include <Shlwapi.h>
#include <shellapi.h>
#include <VersionHelpers.h>

#include <algorithm>
#include <cwchar>
#include <cwctype>
#include <cstring>
#include <iterator>

#pragma comment(lib, "Shlwapi.lib")

#include "OptionsDialog.h"
#include "NativeTabControl.h"
#include "QTTabBarClass.h"
#include "Resource.h"
#include "AliasStoreNative.h"
#include "ClosedTabHistoryStore.h"
#include "GroupsManagerNative.h"
#include "InstanceManagerNative.h"
#include "TextInputDialog.h"

namespace {
std::wstring NormalizeUrlToPath(const std::wstring& url) {
    if(url.empty()) {
        return {};
    }

    if(url.rfind(L"file:", 0) == 0) {
        DWORD required = 0;
        if(PathCreateFromUrlW(url.c_str(), nullptr, &required, 0) == E_POINTER && required > 0) {
            std::wstring buffer(required, L'\0');
            if(SUCCEEDED(PathCreateFromUrlW(url.c_str(), buffer.data(), &required, 0))) {
                buffer.resize(required - 1); // remove trailing null
                return buffer;
            }
        }
    }

    return url;
}

std::wstring JoinTabList(const std::vector<std::wstring>& tabs) {
    std::wstring result;
    for(const auto& path : tabs) {
        if(path.empty()) {
            continue;
        }
        if(!result.empty()) {
            result.append(L";");
        }
        result.append(path);
    }
    return result;
}

std::vector<std::wstring> SplitTabsString(const std::wstring& value) {
    std::vector<std::wstring> output;
    std::wstring token;
    for(wchar_t ch : value) {
        if(ch == L';') {
            if(!token.empty()) {
                output.push_back(token);
                token.clear();
            }
        } else {
            token.push_back(ch);
        }
    }
    if(!token.empty()) {
        output.push_back(token);
    }
    return output;
}

std::wstring ExtractLeafName(const std::wstring& path) {
    if(path.empty()) {
        return {};
    }
    const wchar_t* name = PathFindFileNameW(path.c_str());
    if(name != nullptr && *name != L'\0') {
        return name;
    }
    return path;
}

} // namespace

_ATL_FUNC_INFO TabBarHost::kBeforeNavigate2Info = {CC_STDCALL, VT_EMPTY, 7,
                                                   {VT_DISPATCH, VT_BYREF | VT_VARIANT, VT_BYREF | VT_VARIANT,
                                                    VT_BYREF | VT_VARIANT, VT_BYREF | VT_VARIANT, VT_BYREF | VT_VARIANT,
                                                    VT_BYREF | VT_BOOL}};
_ATL_FUNC_INFO TabBarHost::kNavigateComplete2Info = {CC_STDCALL, VT_EMPTY, 2,
                                                     {VT_DISPATCH, VT_BYREF | VT_VARIANT}};
_ATL_FUNC_INFO TabBarHost::kDocumentCompleteInfo = {CC_STDCALL, VT_EMPTY, 2,
                                                    {VT_DISPATCH, VT_BYREF | VT_VARIANT}};
_ATL_FUNC_INFO TabBarHost::kNavigateErrorInfo = {CC_STDCALL, VT_EMPTY, 4,
                                                 {VT_DISPATCH, VT_BYREF | VT_VARIANT, VT_BYREF | VT_VARIANT,
                                                  VT_BYREF | VT_VARIANT}};
_ATL_FUNC_INFO TabBarHost::kOnQuitInfo = {CC_STDCALL, VT_EMPTY, 0, {}};

TabBarHost::TabBarHost(ITabBarHostOwner& owner) noexcept
    : m_owner(owner)
    , m_browserCookie(0)
    , m_dpiX(96)
    , m_dpiY(96)
    , m_hasFocus(false)
    , m_visible(false) {
}

TabBarHost::~TabBarHost() {
    ATLTRACE(L"TabBarHost::~TabBarHost\n");
    DisconnectBrowserEvents();
}

void TabBarHost::Initialize() {
    RestoreSessionState();
    if(m_spBrowser) {
        ConnectBrowserEvents();
    }
}

void TabBarHost::SetExplorer(IWebBrowser2* browser) {
    if(browser == m_spBrowser.p) {
        return;
    }

    DisconnectBrowserEvents();
    m_spBrowser = browser;
    if(m_spBrowser) {
        ConnectBrowserEvents();
        CComBSTR location;
        if(SUCCEEDED(m_spBrowser->get_LocationURL(&location)) && location != nullptr) {
            UpdateActivePath(NormalizeUrlToPath(std::wstring(location, SysStringLen(location))));
        }
    }
}

void TabBarHost::ClearExplorer() {
    DisconnectBrowserEvents();
    m_spBrowser.Release();
}

void TabBarHost::ExecuteCommand(UINT commandId) {
    ATLTRACE(L"TabBarHost::ExecuteCommand %u\n", commandId);
    std::optional<std::size_t> contextIndex = ResolveContextTabIndex();
    std::size_t activeIndex = m_tabControl ? m_tabControl->GetActiveIndex() : 0;
    std::size_t targetIndex = contextIndex.value_or(activeIndex);
    std::wstring targetPath = GetTabPath(targetIndex);

    if(commandId >= ID_CONTEXT_GROUP_BASE && commandId < ID_CONTEXT_HISTORY_BASE) {
        std::size_t groupIndex = commandId - ID_CONTEXT_GROUP_BASE;
        if(!targetPath.empty()) {
            if(auto group = qttabbar::GroupsManagerNative::Instance().GetGroupByIndex(groupIndex)) {
                qttabbar::GroupsManagerNative::Instance().AppendPaths(group->name, {targetPath});
            }
        }
        m_contextTabIndex.reset();
        return;
    }

    if(commandId >= ID_CONTEXT_HISTORY_BASE && commandId <= ID_CONTEXT_DYNAMIC_LAST) {
        std::size_t historyIndex = commandId - ID_CONTEXT_HISTORY_BASE;
        RestoreClosedTabByIndex(historyIndex);
        m_contextTabIndex.reset();
        return;
    }

    switch(commandId) {
    case ID_CONTEXT_NEWTAB:
        AddTab(m_currentPath, true, true);
        break;
    case ID_CONTEXT_CLOSETAB:
        CloseTabAt(targetIndex);
        break;
    case ID_CONTEXT_CLOSE_LEFT:
        if(m_tabControl) {
            auto closed = m_tabControl->CloseTabsToLeft(targetIndex);
            if(!closed.empty()) {
                for(const auto& path : closed) {
                    RecordClosedEntry(path);
                }
                PersistClosedHistory();
            }
        }
        break;
    case ID_CONTEXT_CLOSE_RIGHT:
        if(m_tabControl) {
            auto closed = m_tabControl->CloseTabsToRight(targetIndex);
            if(!closed.empty()) {
                for(const auto& path : closed) {
                    RecordClosedEntry(path);
                }
                PersistClosedHistory();
            }
        }
        break;
    case ID_CONTEXT_CLOSE_OTHERS:
        if(m_tabControl) {
            auto closed = m_tabControl->CloseAllExcept(targetIndex);
            if(!closed.empty()) {
                for(const auto& path : closed) {
                    RecordClosedEntry(path);
                }
                PersistClosedHistory();
            }
        }
        break;
    case ID_CONTEXT_RESTORELASTCLOSED:
        RestoreLastClosed();
        break;
    case ID_CONTEXT_TOGGLE_LOCK:
        ToggleLockTab(targetIndex);
        break;
    case ID_CONTEXT_CLONE_TAB:
        if(!targetPath.empty()) {
            AddTab(targetPath, true, true);
        }
        break;
    case ID_CONTEXT_CREATE_GROUP:
        if(!targetPath.empty()) {
            std::wstring caption = LoadStringResource(IDS_CONTEXT_DIALOG_GROUP_TITLE, L"Create Group");
            std::wstring prompt = LoadStringResource(IDS_CONTEXT_DIALOG_GROUP_PROMPT, L"Enter a name for the new group:");
            if(auto name = TextInputDialog::Show(m_hWnd, caption, prompt, {}); name && !name->empty()) {
                qttabbar::GroupsManagerNative::Instance().AddGroup(*name, {targetPath});
            }
        }
        break;
    case ID_CONTEXT_OPEN_NEW_WINDOW:
        if(!targetPath.empty()) {
            ::ShellExecuteW(nullptr, L"open", L"explorer.exe", targetPath.c_str(), nullptr, SW_SHOWNORMAL);
        }
        break;
    case ID_CONTEXT_COPY_PATH:
        CopyPathToClipboard(targetPath);
        break;
    case ID_CONTEXT_PROPERTIES:
        ShowProperties(targetPath);
        break;
    case ID_CONTEXT_HISTORY_CLEAR:
        ClearClosedHistory();
        break;
    case ID_CONTEXT_OPEN_COMMAND:
        OpenCommandPrompt(targetPath);
        break;
    case ID_CONTEXT_EDIT_ALIAS: {
        if(targetPath.empty()) {
            break;
        }
        std::wstring caption = LoadStringResource(IDS_CONTEXT_DIALOG_ALIAS_TITLE, L"Tab Alias");
        std::wstring prompt = LoadStringResource(IDS_CONTEXT_DIALOG_ALIAS_PROMPT, L"Enter an alias for this tab:");
        std::wstring initial;
        if(auto existing = qttabbar::AliasStoreNative::Instance().GetAlias(targetPath)) {
            initial = *existing;
        }
        auto alias = TextInputDialog::Show(m_hWnd, caption, prompt, initial);
        if(alias.has_value()) {
            std::wstring trimmed = *alias;
            trimmed.erase(trimmed.begin(), std::find_if(trimmed.begin(), trimmed.end(), [](wchar_t ch) { return !iswspace(ch); }));
            while(!trimmed.empty() && iswspace(trimmed.back())) {
                trimmed.pop_back();
            }
            SetTabAlias(targetIndex, trimmed);
            if(trimmed.empty()) {
                qttabbar::AliasStoreNative::Instance().ClearAlias(targetPath);
            } else {
                qttabbar::AliasStoreNative::Instance().SetAlias(targetPath, trimmed);
            }
        }
        break;
    }
    case ID_CONTEXT_REFRESH:
        RefreshExplorer();
        break;
    case ID_CONTEXT_OPTIONS:
        OpenOptions();
        break;
    default:
        break;
    }

    m_contextTabIndex.reset();
}

void TabBarHost::ShowContextMenu(const POINT& screenPoint) {
    if(!m_tabControl) {
        return;
    }

    auto targetIndexOpt = ResolveContextTabIndex();
    std::size_t activeIndex = m_tabControl->GetActiveIndex();
    std::size_t targetIndex = targetIndexOpt.value_or(activeIndex);
    std::size_t tabCount = m_tabControl->GetCount();
    if(tabCount == 0) {
        targetIndex = 0;
    } else if(targetIndex >= tabCount) {
        targetIndex = tabCount - 1;
    }
    std::wstring targetPath = GetTabPath(targetIndex);

    HMENU menu = BuildContextMenu(targetIndex, targetPath);
    if(menu == nullptr) {
        return;
    }

    ::TrackPopupMenuEx(menu, TPM_LEFTALIGN | TPM_LEFTBUTTON, screenPoint.x, screenPoint.y, m_hWnd, nullptr);
    ::DestroyMenu(menu);
}

void TabBarHost::SaveSessionState() const {
    std::vector<std::wstring> paths = m_tabControl ? m_tabControl->GetTabPaths() : std::vector<std::wstring>();
    std::wstring serialized = JoinTabList(paths);
    ATLTRACE(L"TabBarHost::SaveSessionState tabs='%s'\n", serialized.c_str());

    CRegKey key;
    if(key.Create(HKEY_CURRENT_USER, kRegistryRoot) == ERROR_SUCCESS) {
        key.SetStringValue(kTabsValueName, serialized.c_str());
    }
}

void TabBarHost::RestoreSessionState() {
    CRegKey key;
    if(key.Open(HKEY_CURRENT_USER, kRegistryRoot, KEY_READ) != ERROR_SUCCESS) {
        return;
    }

    ULONG chars = 0;
    if(key.QueryStringValue(kTabsValueName, nullptr, &chars) != ERROR_SUCCESS || chars == 0) {
        return;
    }

    std::wstring buffer(chars, L'\0');
    if(key.QueryStringValue(kTabsValueName, buffer.data(), &chars) == ERROR_SUCCESS) {
        if(!buffer.empty() && buffer.back() == L'\0') {
            buffer.pop_back();
        }
        auto restored = SplitTabsString(buffer);
        if(m_tabControl) {
            for(const auto& tabPath : restored) {
                m_tabControl->AddTab(tabPath, false, true);
            }
            if(m_tabControl->GetCount() > 0) {
                auto path = m_tabControl->ActivateTab(0);
                m_currentPath = path;
            }
        }
        LogTabsState(L"RestoreSessionState");
    }
}

void TabBarHost::OnBandVisibilityChanged(bool visible) {
    m_visible = visible;
    ATLTRACE(L"TabBarHost::OnBandVisibilityChanged visible=%d\n", visible);
    if(visible) {
        StartTimers();
    } else {
        StopTimers();
    }
}

bool TabBarHost::HandleAccelerator(MSG* pMsg) {
    if(pMsg == nullptr) {
        return false;
    }
    if(pMsg->message == WM_KEYDOWN || pMsg->message == WM_SYSKEYDOWN) {
        switch(pMsg->wParam) {
        case VK_LEFT:
            ActivatePreviousTab();
            return true;
        case VK_RIGHT:
            ActivateNextTab();
            return true;
        case VK_DELETE:
            CloseActiveTab();
            return true;
        case 'T':
            if(::GetKeyState(VK_CONTROL) < 0) {
                AddTab(m_currentPath, true, true);
                return true;
            }
            break;
        case 'W':
            if(::GetKeyState(VK_CONTROL) < 0) {
                CloseActiveTab();
                return true;
            }
            break;
        case 'Z':
            if((::GetKeyState(VK_CONTROL) < 0) && (::GetKeyState(VK_SHIFT) < 0)) {
                RestoreLastClosed();
                return true;
            }
            break;
        default:
            break;
        }
    }
    return false;
}

void TabBarHost::OnParentDestroyed() {
    StopTimers();
    SaveSessionState();
    DisconnectBrowserEvents();
}

LRESULT TabBarHost::OnCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    ATLTRACE(L"TabBarHost::OnCreate\n");
    if(!m_tabControl) {
        m_tabControl = std::make_unique<NativeTabControl>(*this);
    }
    if(m_tabControl && !m_tabControl->IsWindow()) {
        RECT rc{};
        ::GetClientRect(m_hWnd, &rc);
        HWND hwnd = m_tabControl->Create(m_hWnd, rc, L"",
                                         WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS);
        if(hwnd == nullptr) {
            return -1;
        }
    }
    LoadClosedHistory();
    StartTimers();
#if defined(_WIN32_WINNT) && _WIN32_WINNT >= _WIN32_WINNT_WIN10
    if(::IsWindows10OrGreater()) {
        m_dpiX = ::GetDpiForWindow(m_hWnd);
        m_dpiY = m_dpiX;
    }
#endif
    return 0;
}

LRESULT TabBarHost::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    ATLTRACE(L"TabBarHost::OnDestroy\n");
    StopTimers();
    SaveSessionState();
    DisconnectBrowserEvents();
    if(m_tabControl) {
        if(m_tabControl->IsWindow()) {
            m_tabControl->DestroyWindow();
        }
        m_tabControl.reset();
    }
    return 0;
}

LRESULT TabBarHost::OnSize(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    int width = LOWORD(lParam);
    int height = HIWORD(lParam);
    if(m_tabControl && m_tabControl->IsWindow()) {
        ::MoveWindow(m_tabControl->m_hWnd, 0, 0, width, height, TRUE);
        m_tabControl->EnsureLayout();
    }
    return 0;
}

LRESULT TabBarHost::OnTimer(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    if(wParam == ID_TIMER_SELECTTAB) {
        ATLTRACE(L"TabBarHost::OnTimer selectTab\n");
    } else if(wParam == ID_TIMER_CONTEXTMENU) {
        ATLTRACE(L"TabBarHost::OnTimer contextMenu\n");
    }
    return 0;
}

LRESULT TabBarHost::OnContextMenu(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    POINT pt{};
    if(reinterpret_cast<HWND>(wParam) == m_hWnd) {
        pt.x = GET_X_LPARAM(lParam);
        pt.y = GET_Y_LPARAM(lParam);
        if(pt.x == -1 && pt.y == -1) {
            RECT rc{};
            ::GetClientRect(m_hWnd, &rc);
            pt.x = rc.left;
            pt.y = rc.bottom;
            ::ClientToScreen(m_hWnd, &pt);
        }
    } else {
        ::GetCursorPos(&pt);
    }
    m_contextTabIndex.reset();
    ShowContextMenu(pt);
    return 0;
}

LRESULT TabBarHost::OnCommand(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    ExecuteCommand(LOWORD(wParam));
    return 0;
}

LRESULT TabBarHost::OnSetFocus(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    m_hasFocus = true;
    ATLTRACE(L"TabBarHost::OnSetFocus\n");
    if(m_tabControl && m_tabControl->IsWindow()) {
        ::SetFocus(m_tabControl->m_hWnd);
    }
    m_owner.NotifyTabHostFocusChange(TRUE);
    return 0;
}

LRESULT TabBarHost::OnKillFocus(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    m_hasFocus = false;
    ATLTRACE(L"TabBarHost::OnKillFocus\n");
    m_owner.NotifyTabHostFocusChange(FALSE);
    return 0;
}

LRESULT TabBarHost::OnDpiChanged(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    UINT newDpiX = LOWORD(wParam);
    UINT newDpiY = HIWORD(wParam);
    ATLTRACE(L"TabBarHost::OnDpiChanged %u %u\n", newDpiX, newDpiY);
    m_dpiX = newDpiX;
    m_dpiY = newDpiY;
    if(lParam != 0) {
        const RECT* suggested = reinterpret_cast<RECT*>(lParam);
        ::SetWindowPos(m_hWnd, nullptr, suggested->left, suggested->top, suggested->right - suggested->left,
                       suggested->bottom - suggested->top, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    return 0;
}

LRESULT TabBarHost::OnGetDlgCode(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    return DLGC_WANTARROWS | DLGC_WANTTAB | DLGC_WANTCHARS;
}

LRESULT TabBarHost::OnKeyDown(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    MSG msg{};
    msg.message = WM_KEYDOWN;
    msg.wParam = wParam;
    if(!HandleAccelerator(&msg)) {
        bHandled = FALSE;
    }
    return 0;
}

void __stdcall TabBarHost::OnBeforeNavigate2(IDispatch* /*pDisp*/, VARIANT* url, VARIANT* /*flags*/, VARIANT* /*targetFrameName*/,
                                             VARIANT* /*postData*/, VARIANT* /*headers*/, VARIANT_BOOL* /*cancel*/) {
    std::wstring path = NormalizeUrlToPath(VariantToString(url));
    ATLTRACE(L"TabBarHost::OnBeforeNavigate2 %s\n", path.c_str());
    if(!path.empty()) {
        m_pendingNavigation = path;
    }
}

void __stdcall TabBarHost::OnNavigateComplete2(IDispatch* /*pDisp*/, VARIANT* url) {
    std::wstring path = NormalizeUrlToPath(VariantToString(url));
    ATLTRACE(L"TabBarHost::OnNavigateComplete2 %s\n", path.c_str());
    UpdateActivePath(path);
}

void __stdcall TabBarHost::OnDocumentComplete(IDispatch* /*pDisp*/, VARIANT* url) {
    std::wstring path = NormalizeUrlToPath(VariantToString(url));
    ATLTRACE(L"TabBarHost::OnDocumentComplete %s\n", path.c_str());
    UpdateActivePath(path);
}

void __stdcall TabBarHost::OnNavigateError(IDispatch* /*pDisp*/, VARIANT* url, VARIANT* /*targetFrameName*/, VARIANT* statusCode,
                                           VARIANT_BOOL* /*cancel*/) {
    std::wstring path = NormalizeUrlToPath(VariantToString(url));
    std::wstring status = VariantToString(statusCode);
    ATLTRACE(L"TabBarHost::OnNavigateError %s status=%s\n", path.c_str(), status.c_str());
}

void __stdcall TabBarHost::OnQuit() {
    ATLTRACE(L"TabBarHost::OnQuit\n");
    SaveSessionState();
}

void TabBarHost::ConnectBrowserEvents() {
    if(!m_spBrowser || m_browserCookie != 0) {
        return;
    }
    ATLTRACE(L"TabBarHost::ConnectBrowserEvents\n");
    AtlAdvise(m_spBrowser, static_cast<IDispatch*>(this), DIID_DWebBrowserEvents2, &m_browserCookie);
}

void TabBarHost::DisconnectBrowserEvents() {
    if(m_spBrowser && m_browserCookie != 0) {
        ATLTRACE(L"TabBarHost::DisconnectBrowserEvents\n");
        AtlUnadvise(m_spBrowser, DIID_DWebBrowserEvents2, m_browserCookie);
        m_browserCookie = 0;
    }
}

void TabBarHost::StartTimers() {
    if(!m_visible) {
        m_visible = true;
    }
    if(m_hWnd != nullptr) {
        ::SetTimer(m_hWnd, ID_TIMER_SELECTTAB, kSelectTabTimerMs, nullptr);
        ::SetTimer(m_hWnd, ID_TIMER_CONTEXTMENU, kContextMenuTimerMs, nullptr);
    }
}

void TabBarHost::StopTimers() {
    if(m_hWnd != nullptr) {
        ::KillTimer(m_hWnd, ID_TIMER_SELECTTAB);
        ::KillTimer(m_hWnd, ID_TIMER_CONTEXTMENU);
    }
}

std::wstring TabBarHost::VariantToString(const VARIANT* value) const {
    if(value == nullptr) {
        return {};
    }
    CComVariant normalized;
    if(FAILED(normalized.Copy(value))) {
        return {};
    }
    if(FAILED(normalized.ChangeType(VT_BSTR))) {
        return {};
    }
    if(normalized.bstrVal == nullptr) {
        return {};
    }
    return std::wstring(normalized.bstrVal, SysStringLen(normalized.bstrVal));
}

void TabBarHost::UpdateActivePath(const std::wstring& path) {
    if(path.empty()) {
        return;
    }
    std::wstring effectivePath = path;
    if(m_pendingNavigation.has_value()) {
        effectivePath = *m_pendingNavigation;
        m_pendingNavigation.reset();
    }

    m_currentPath = effectivePath;
    AddTab(effectivePath, true, false);
    LogTabsState(L"UpdateActivePath");
}

void TabBarHost::AddTab(const std::wstring& path, bool makeActive, bool allowDuplicate) {
    if(path.empty()) {
        return;
    }
    if(m_tabControl) {
        std::size_t index = m_tabControl->AddTab(path, makeActive, allowDuplicate);
        if(auto alias = qttabbar::AliasStoreNative::Instance().GetAlias(path)) {
            m_tabControl->SetAlias(index, *alias);
        }
    }
    if(makeActive) {
        m_currentPath = path;
    }
}

void TabBarHost::ActivateTab(std::size_t index) {
    if(!m_tabControl) {
        return;
    }
    std::wstring path = m_tabControl->ActivateTab(index);
    if(!path.empty()) {
        m_currentPath = path;
    }
    LogTabsState(L"ActivateTab");
}

void TabBarHost::ActivateNextTab() {
    if(!m_tabControl || m_tabControl->GetCount() == 0) {
        return;
    }
    std::wstring path = m_tabControl->ActivateNextTab();
    if(!path.empty()) {
        m_currentPath = path;
    }
}

void TabBarHost::ActivatePreviousTab() {
    if(!m_tabControl || m_tabControl->GetCount() == 0) {
        return;
    }
    std::wstring path = m_tabControl->ActivatePreviousTab();
    if(!path.empty()) {
        m_currentPath = path;
    }
}

void TabBarHost::CloseActiveTab() {
    if(!m_tabControl || m_tabControl->GetCount() == 0) {
        return;
    }
    CloseTabAt(m_tabControl->GetActiveIndex());
}

void TabBarHost::RestoreLastClosed() {
    if(m_closedHistory.empty()) {
        return;
    }
    std::wstring path = m_closedHistory.front();
    m_closedHistory.pop_front();
    ATLTRACE(L"TabBarHost::RestoreLastClosed %s\n", path.c_str());
    AddTab(path, true, true);
    PersistClosedHistory();
}

void TabBarHost::CloseTabAt(std::size_t index) {
    if(!m_tabControl || index >= m_tabControl->GetCount()) {
        return;
    }
    auto closed = m_tabControl->CloseTab(index);
    if(!closed) {
        return;
    }
    RecordClosedEntry(*closed);
    if(auto active = m_tabControl->GetActivePath()) {
        m_currentPath = *active;
    } else {
        m_currentPath.clear();
    }
    PersistClosedHistory();
    LogTabsState(L"CloseTabAt");
}

void TabBarHost::RefreshExplorer() {
    if(m_spBrowser) {
        ATLTRACE(L"TabBarHost::RefreshExplorer\n");
        m_spBrowser->Refresh();
    }
}

void TabBarHost::OpenOptions() {
    ATLTRACE(L"TabBarHost::OpenOptions\n");
    HWND ownerHwnd = m_owner.GetHostRebarWindow();
    if(ownerHwnd == nullptr) {
        ownerHwnd = m_owner.GetHostWindow();
    }
    qttabbar::OptionsDialog::Open(ownerHwnd);
}

std::vector<std::wstring> TabBarHost::GetOpenTabs() const {
    if(!m_tabControl) {
        return {};
    }
    return m_tabControl->GetTabPaths();
}

std::vector<std::wstring> TabBarHost::GetClosedTabHistory() const {
    return std::vector<std::wstring>(m_closedHistory.begin(), m_closedHistory.end());
}

void TabBarHost::ActivateTabByIndex(std::size_t index) {
    ActivateTab(index);
}

void TabBarHost::RestoreClosedTabByIndex(std::size_t index) {
    if(index >= m_closedHistory.size()) {
        return;
    }
    auto it = m_closedHistory.begin();
    std::advance(it, static_cast<std::ptrdiff_t>(index));
    std::wstring path = *it;
    m_closedHistory.erase(it);
    AddTab(path, true, true);
    PersistClosedHistory();
}

void TabBarHost::CloneActiveTab() {
    if(m_currentPath.empty()) {
        return;
    }
    AddTab(m_currentPath, true, true);
}

void TabBarHost::CloseAllTabsExceptActive() {
    if(!m_tabControl || m_tabControl->GetCount() == 0) {
        return;
    }
    std::size_t activeIndex = m_tabControl->GetActiveIndex();
    auto closed = m_tabControl->CloseAllExcept(activeIndex);
    for(const auto& path : closed) {
        RecordClosedEntry(path);
    }
    if(!closed.empty()) {
        PersistClosedHistory();
    }
    if(auto active = m_tabControl->GetActivePath()) {
        m_currentPath = *active;
    }
    LogTabsState(L"CloseAllTabsExceptActive");
}

void TabBarHost::CloseTabsToLeft() {
    if(!m_tabControl || m_tabControl->GetCount() == 0) {
        return;
    }
    std::size_t activeIndex = m_tabControl->GetActiveIndex();
    auto closed = m_tabControl->CloseTabsToLeft(activeIndex);
    for(const auto& path : closed) {
        RecordClosedEntry(path);
    }
    if(!closed.empty()) {
        PersistClosedHistory();
    }
    LogTabsState(L"CloseTabsToLeft");
}

void TabBarHost::CloseTabsToRight() {
    if(!m_tabControl || m_tabControl->GetCount() == 0) {
        return;
    }
    std::size_t activeIndex = m_tabControl->GetActiveIndex();
    auto closed = m_tabControl->CloseTabsToRight(activeIndex);
    for(const auto& path : closed) {
        RecordClosedEntry(path);
    }
    if(!closed.empty()) {
        PersistClosedHistory();
    }
    LogTabsState(L"CloseTabsToRight");
}

void TabBarHost::GoUpOneLevel() {
    if(m_currentPath.empty()) {
        return;
    }
    std::wstring path = m_currentPath;
    path.push_back(L'\0');
    if(PathRemoveFileSpecW(path.data())) {
        path.resize(wcslen(path.c_str()));
        AddTab(path, true, true);
    }
}

void TabBarHost::NavigateBack() {
    if(m_spBrowser) {
        m_spBrowser->GoBack();
    }
}

void TabBarHost::NavigateForward() {
    if(m_spBrowser) {
        m_spBrowser->GoForward();
    }
}

bool TabBarHost::OpenCapturedWindow(const std::wstring& path) {
    if(path.empty()) {
        return false;
    }
    AddTab(path, true, true);
    LogTabsState(L"OpenCapturedWindow");
    return true;
}

void TabBarHost::OpenGroupByIndex(std::size_t index) {
    auto group = qttabbar::GroupsManagerNative::Instance().GetGroupByIndex(index);
    if(!group.has_value() || group->paths.empty()) {
        return;
    }
    bool first = true;
    for(const auto& path : group->paths) {
        if(path.empty()) {
            continue;
        }
        AddTab(path, first, true);
        first = false;
    }
    LogTabsState(L"OpenGroupByIndex");
}

void TabBarHost::TrimClosedHistory() {
    while(m_closedHistory.size() > kMaxClosedHistory) {
        m_closedHistory.pop_back();
    }
}

void TabBarHost::LoadClosedHistory() {
    m_closedHistory = qttabbar::ClosedTabHistoryStore::Instance().Load();
    TrimClosedHistory();
}

void TabBarHost::PersistClosedHistory() const {
    qttabbar::ClosedTabHistoryStore::Instance().Save(m_closedHistory);
}

void TabBarHost::ClearClosedHistory() {
    m_closedHistory.clear();
    qttabbar::ClosedTabHistoryStore::Instance().Clear();
}

void TabBarHost::RecordClosedEntry(const std::wstring& path) {
    if(path.empty()) {
        return;
    }
    auto it = std::remove(m_closedHistory.begin(), m_closedHistory.end(), path);
    if(it != m_closedHistory.end()) {
        m_closedHistory.erase(it, m_closedHistory.end());
    }
    m_closedHistory.push_front(path);
    TrimClosedHistory();
}

std::optional<std::size_t> TabBarHost::ResolveContextTabIndex() const {
    if(!m_tabControl) {
        return std::nullopt;
    }
    if(m_contextTabIndex.has_value() && *m_contextTabIndex < m_tabControl->GetCount()) {
        return m_contextTabIndex;
    }
    return std::nullopt;
}

std::wstring TabBarHost::GetTabPath(std::size_t index) const {
    if(!m_tabControl) {
        return {};
    }
    return m_tabControl->GetPath(index);
}

bool TabBarHost::IsTabLocked(std::size_t index) const {
    if(!m_tabControl) {
        return false;
    }
    return m_tabControl->IsLocked(index);
}

void TabBarHost::ToggleLockTab(std::size_t index) {
    if(!m_tabControl) {
        return;
    }
    bool locked = m_tabControl->IsLocked(index);
    m_tabControl->SetLocked(index, !locked);
}

void TabBarHost::SetTabAlias(std::size_t index, const std::wstring& alias) {
    if(!m_tabControl) {
        return;
    }
    m_tabControl->SetAlias(index, alias);
}

void TabBarHost::CopyPathToClipboard(const std::wstring& path) const {
    if(path.empty()) {
        return;
    }
    if(!::OpenClipboard(m_hWnd)) {
        return;
    }
    ::EmptyClipboard();
    SIZE_T bytes = (path.size() + 1) * sizeof(wchar_t);
    HGLOBAL handle = ::GlobalAlloc(GMEM_MOVEABLE, bytes);
    if(handle != nullptr) {
        void* data = ::GlobalLock(handle);
        if(data != nullptr) {
            std::memcpy(data, path.c_str(), bytes);
            ::GlobalUnlock(handle);
            ::SetClipboardData(CF_UNICODETEXT, handle);
        } else {
            ::GlobalFree(handle);
        }
    }
    ::CloseClipboard();
}

void TabBarHost::OpenCommandPrompt(const std::wstring& path) const {
    std::wstring directory = path.empty() ? m_currentPath : path;
    if(directory.empty()) {
        return;
    }
    ::ShellExecuteW(nullptr, L"open", L"cmd.exe", nullptr, directory.c_str(), SW_SHOWNORMAL);
}

void TabBarHost::ShowProperties(const std::wstring& path) const {
    if(path.empty()) {
        return;
    }
    SHELLEXECUTEINFOW exec{};
    exec.cbSize = sizeof(exec);
    exec.fMask = SEE_MASK_INVOKEIDLIST;
    exec.lpVerb = L"properties";
    exec.lpFile = path.c_str();
    exec.nShow = SW_SHOWNORMAL;
    ::ShellExecuteExW(&exec);
}

std::wstring TabBarHost::LoadStringResource(UINT id, const wchar_t* fallback) const {
    wchar_t buffer[256] = {};
    HINSTANCE instance = _AtlBaseModule.GetResourceInstance();
    int length = ::LoadStringW(instance, id, buffer, static_cast<int>(_countof(buffer)));
    if(length > 0) {
        return std::wstring(buffer, length);
    }
    return fallback != nullptr ? std::wstring(fallback) : std::wstring();
}

HMENU TabBarHost::BuildContextMenu(std::size_t targetIndex, std::wstring targetPath) {
    HMENU menu = ::CreatePopupMenu();
    if(menu == nullptr) {
        return nullptr;
    }

    std::size_t tabCount = m_tabControl ? m_tabControl->GetCount() : 0;
    if(tabCount == 0) {
        targetIndex = 0;
    } else if(targetIndex >= tabCount) {
        targetIndex = tabCount - 1;
    }

    bool hasTabs = m_tabControl && tabCount > 0;
    if(hasTabs) {
        targetPath = GetTabPath(targetIndex);
    } else {
        targetPath.clear();
    }
    bool locked = hasTabs && IsTabLocked(targetIndex);
    bool canCloseTab = hasTabs && m_tabControl->CanCloseTab(targetIndex);
    bool canCloseLeft = hasTabs && m_tabControl->HasClosableTabsToLeft(targetIndex);
    bool canCloseRight = hasTabs && m_tabControl->HasClosableTabsToRight(targetIndex);
    bool canCloseOthers = hasTabs && m_tabControl->HasClosableOtherTabs(targetIndex);
    bool hasHistory = !m_closedHistory.empty();

    const std::wstring newTabText = LoadStringResource(IDS_CONTEXT_MENU_NEW_TAB, L"&New Tab");
    const std::wstring closeTabText = LoadStringResource(IDS_CONTEXT_MENU_CLOSE_TAB, L"&Close Tab");
    const std::wstring closeRightText = LoadStringResource(IDS_CONTEXT_MENU_CLOSE_RIGHT, L"Close Tabs to the &Right");
    const std::wstring closeLeftText = LoadStringResource(IDS_CONTEXT_MENU_CLOSE_LEFT, L"Close Tabs to the &Left");
    const std::wstring closeOthersText = LoadStringResource(IDS_CONTEXT_MENU_CLOSE_OTHERS, L"Close &Other Tabs");
    const std::wstring restoreLastText = LoadStringResource(IDS_CONTEXT_MENU_RESTORE_LAST, L"Restore &Last Closed");
    const std::wstring groupCaption = LoadStringResource(IDS_CONTEXT_MENU_GROUP_CAPTION, L"Add to &Group");
    const std::wstring createGroupText = LoadStringResource(IDS_CONTEXT_MENU_CREATE_GROUP, L"Create &Group...");
    const std::wstring cloneTabText = LoadStringResource(IDS_CONTEXT_MENU_CLONE_TAB, L"&Clone Tab");
    const std::wstring openNewWindowText = LoadStringResource(IDS_CONTEXT_MENU_OPEN_NEW_WIN, L"Open in New &Window");
    const std::wstring copyPathText = LoadStringResource(IDS_CONTEXT_MENU_COPY_PATH, L"&Copy Path");
    const std::wstring propertiesText = LoadStringResource(IDS_CONTEXT_MENU_PROPERTIES, L"&Properties");
    const std::wstring historyText = LoadStringResource(IDS_CONTEXT_MENU_HISTORY, L"&History");
    const std::wstring commandPromptText = LoadStringResource(IDS_CONTEXT_MENU_COMMAND, L"Open &Command Prompt Here");
    const std::wstring editAliasText = LoadStringResource(IDS_CONTEXT_MENU_EDIT_ALIAS, L"Edit &Alias...");
    const std::wstring refreshText = LoadStringResource(IDS_CONTEXT_MENU_REFRESH, L"&Refresh");
    const std::wstring optionsText = LoadStringResource(IDS_CONTEXT_MENU_OPTIONS, L"&Options...");

    auto appendItem = [&](UINT id, const std::wstring& text, bool enabled = true) {
        ::AppendMenuW(menu, (enabled ? MF_STRING : MF_STRING | MF_GRAYED), id, text.c_str());
    };
    auto appendPopup = [&](HMENU subMenu, const std::wstring& text, bool enabled) {
        ::AppendMenuW(menu, (enabled ? MF_POPUP | MF_STRING : MF_POPUP | MF_STRING | MF_GRAYED),
                      reinterpret_cast<UINT_PTR>(subMenu), text.c_str());
    };

    appendItem(ID_CONTEXT_NEWTAB, newTabText, true);
    appendItem(ID_CONTEXT_CLOSETAB, closeTabText, canCloseTab);
    appendItem(ID_CONTEXT_CLOSE_RIGHT, closeRightText, canCloseRight);
    appendItem(ID_CONTEXT_CLOSE_LEFT, closeLeftText, canCloseLeft);
    appendItem(ID_CONTEXT_CLOSE_OTHERS, closeOthersText, canCloseOthers);
    appendItem(ID_CONTEXT_RESTORELASTCLOSED, restoreLastText, hasHistory);
    ::AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);

    HMENU groupMenu = ::CreatePopupMenu();
    if(groupMenu) {
        bool hasGroupEntries = PopulateGroupMenu(groupMenu, targetPath);
        appendPopup(groupMenu, groupCaption, hasGroupEntries && !targetPath.empty());
    }

    appendItem(ID_CONTEXT_CREATE_GROUP, createGroupText, !targetPath.empty());
    ::AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);

    const std::wstring lockText = locked ? LoadStringResource(IDS_CONTEXT_MENU_UNLOCK_TAB, L"&Unlock Tab")
                                         : LoadStringResource(IDS_CONTEXT_MENU_LOCK_TAB, L"&Lock Tab");
    appendItem(ID_CONTEXT_TOGGLE_LOCK, lockText, hasTabs);
    appendItem(ID_CONTEXT_CLONE_TAB, cloneTabText, !targetPath.empty());
    appendItem(ID_CONTEXT_OPEN_NEW_WINDOW, openNewWindowText, !targetPath.empty());
    appendItem(ID_CONTEXT_COPY_PATH, copyPathText, !targetPath.empty());
    appendItem(ID_CONTEXT_PROPERTIES, propertiesText, !targetPath.empty());

    HMENU historyMenu = ::CreatePopupMenu();
    if(historyMenu) {
        PopulateHistoryMenu(historyMenu);
        appendPopup(historyMenu, historyText, hasHistory);
    }

    appendItem(ID_CONTEXT_OPEN_COMMAND, commandPromptText, !targetPath.empty());
    appendItem(ID_CONTEXT_EDIT_ALIAS, editAliasText, hasTabs && !targetPath.empty());
    ::AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    appendItem(ID_CONTEXT_REFRESH, refreshText, true);
    appendItem(ID_CONTEXT_OPTIONS, optionsText, true);

    return menu;
}

bool TabBarHost::PopulateGroupMenu(HMENU menu, const std::wstring& targetPath) const {
    if(menu == nullptr) {
        return false;
    }
    if(targetPath.empty()) {
        return false;
    }
    auto groups = qttabbar::GroupsManagerNative::Instance().GetGroups();
    if(groups.empty()) {
        const std::wstring noGroups = LoadStringResource(IDS_CONTEXT_MENU_NO_GROUPS, L"(no groups)");
        ::AppendMenuW(menu, MF_GRAYED | MF_STRING, 0, noGroups.c_str());
        return false;
    }
    for(std::size_t index = 0; index < groups.size(); ++index) {
        ::AppendMenuW(menu, MF_STRING, ID_CONTEXT_GROUP_BASE + static_cast<UINT>(index), groups[index].name.c_str());
    }
    return true;
}

void TabBarHost::PopulateHistoryMenu(HMENU menu) const {
    if(menu == nullptr) {
        return;
    }
    if(m_closedHistory.empty()) {
        const std::wstring emptyText = LoadStringResource(IDS_CONTEXT_MENU_HISTORY_EMPTY, L"(empty)");
        ::AppendMenuW(menu, MF_GRAYED | MF_STRING, 0, emptyText.c_str());
        return;
    }
    for(std::size_t index = 0; index < m_closedHistory.size(); ++index) {
        std::wstring display = ExtractLeafName(m_closedHistory[index]);
        if(display.empty()) {
            display = m_closedHistory[index];
        }
        ::AppendMenuW(menu, MF_STRING, ID_CONTEXT_HISTORY_BASE + static_cast<UINT>(index), display.c_str());
    }
    ::AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    const std::wstring clearText = LoadStringResource(IDS_CONTEXT_MENU_HISTORY_CLEAR, L"&Clear History");
    ::AppendMenuW(menu, MF_STRING, ID_CONTEXT_HISTORY_CLEAR, clearText.c_str());
}

void TabBarHost::LogTabsState(const wchar_t* source) const {
    if(!m_tabControl) {
        return;
    }
    std::wstring state = JoinTabList(m_tabControl->GetTabPaths());
    ATLTRACE(L"TabBarHost::LogTabsState %s tabs='%s'\n", source, state.c_str());
}

void TabBarHost::OnTabControlTabSelected(std::size_t index) {
    ATLTRACE(L"TabBarHost::OnTabControlTabSelected index=%u\n", static_cast<UINT>(index));
    ActivateTab(index);
}

void TabBarHost::OnTabControlCloseRequested(std::size_t index) {
    ATLTRACE(L"TabBarHost::OnTabControlCloseRequested index=%u\n", static_cast<UINT>(index));
    CloseTabAt(index);
}

void TabBarHost::OnTabControlContextMenuRequested(std::optional<std::size_t> index, const POINT& screenPoint) {
    ATLTRACE(L"TabBarHost::OnTabControlContextMenuRequested (%ld,%ld)\n", screenPoint.x, screenPoint.y);
    m_contextTabIndex = index;
    ShowContextMenu(screenPoint);
}

void TabBarHost::OnTabControlNewTabRequested() {
    ATLTRACE(L"TabBarHost::OnTabControlNewTabRequested\n");
    AddTab(m_currentPath, true, true);
}

void TabBarHost::OnTabControlBeginDrag(std::size_t index, const POINT& screenPoint) {
    ATLTRACE(L"TabBarHost::OnTabControlBeginDrag index=%u (%ld,%ld)\n", static_cast<UINT>(index), screenPoint.x, screenPoint.y);
}
