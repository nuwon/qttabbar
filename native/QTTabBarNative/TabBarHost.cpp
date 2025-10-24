#include "pch.h"
#include "TabBarHost.h"

#include <Shlwapi.h>
#include <VersionHelpers.h>

#include <algorithm>
#include <cwchar>
#include <iterator>

#pragma comment(lib, "Shlwapi.lib")

#include "OptionsDialog.h"
#include "NativeTabControl.h"
#include "QTTabBarClass.h"
#include "Resource.h"

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
    , m_hContextMenu(nullptr)
    , m_dpiX(96)
    , m_dpiY(96)
    , m_hasFocus(false)
    , m_visible(false) {
}

TabBarHost::~TabBarHost() {
    ATLTRACE(L"TabBarHost::~TabBarHost\n");
    DisconnectBrowserEvents();
    DestroyContextMenu();
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
    switch(commandId) {
    case ID_CONTEXT_NEWTAB:
        AddTab(m_currentPath, true, true);
        break;
    case ID_CONTEXT_CLOSETAB:
        CloseActiveTab();
        break;
    case ID_CONTEXT_RESTORELASTCLOSED:
        RestoreLastClosed();
        break;
    case ID_CONTEXT_REFRESH:
        RefreshExplorer();
        break;
    case ID_CONTEXT_OPTIONS:
        OpenOptions();
        break;
    default:
        break;
    }
}

void TabBarHost::ShowContextMenu(const POINT& screenPoint) {
    EnsureContextMenu();
    if(m_hContextMenu == nullptr) {
        return;
    }

    HMENU subMenu = ::GetSubMenu(m_hContextMenu, 0);
    if(subMenu == nullptr) {
        return;
    }

    ::TrackPopupMenuEx(subMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON, screenPoint.x, screenPoint.y, m_hWnd, nullptr);
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
    EnsureContextMenu();
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
    DestroyContextMenu();
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

void TabBarHost::EnsureContextMenu() {
    if(m_hContextMenu != nullptr) {
        return;
    }
    HINSTANCE module = _AtlBaseModule.GetModuleInstance();
    HMENU menu = ::LoadMenuW(module, MAKEINTRESOURCEW(IDR_QTTABBAR_CONTEXT_MENU));
    if(menu != nullptr) {
        m_hContextMenu = menu;
    }
}

void TabBarHost::DestroyContextMenu() {
    if(m_hContextMenu != nullptr) {
        ::DestroyMenu(m_hContextMenu);
        m_hContextMenu = nullptr;
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
        m_tabControl->AddTab(path, makeActive, allowDuplicate);
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
    auto closed = m_tabControl->CloseActiveTab();
    if(closed) {
        ATLTRACE(L"TabBarHost::CloseActiveTab %s\n", closed->c_str());
        m_closedHistory.push_front(*closed);
        TrimClosedHistory();
        if(auto active = m_tabControl->GetActivePath()) {
            m_currentPath = *active;
        } else {
            m_currentPath.clear();
        }
        LogTabsState(L"CloseActiveTab");
    }
}

void TabBarHost::RestoreLastClosed() {
    if(m_closedHistory.empty()) {
        return;
    }
    std::wstring path = m_closedHistory.front();
    m_closedHistory.pop_front();
    ATLTRACE(L"TabBarHost::RestoreLastClosed %s\n", path.c_str());
    AddTab(path, true, true);
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
    m_tabControl->CloseAllExcept(activeIndex, m_closedHistory);
    TrimClosedHistory();
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
    m_tabControl->CloseTabsToLeft(activeIndex, m_closedHistory);
    TrimClosedHistory();
    LogTabsState(L"CloseTabsToLeft");
}

void TabBarHost::CloseTabsToRight() {
    if(!m_tabControl || m_tabControl->GetCount() == 0) {
        return;
    }
    std::size_t activeIndex = m_tabControl->GetActiveIndex();
    m_tabControl->CloseTabsToRight(activeIndex, m_closedHistory);
    TrimClosedHistory();
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

void TabBarHost::TrimClosedHistory() {
    while(m_closedHistory.size() > kMaxClosedHistory) {
        m_closedHistory.pop_back();
    }
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
    if(!m_tabControl) {
        return;
    }

    ATLTRACE(L"TabBarHost::OnTabControlCloseRequested index=%u\n", static_cast<UINT>(index));
    auto closed = m_tabControl->CloseTab(index);
    if(!closed) {
        return;
    }

    m_closedHistory.push_front(*closed);
    TrimClosedHistory();

    if(auto active = m_tabControl->GetActivePath()) {
        m_currentPath = *active;
    } else {
        m_currentPath.clear();
    }

    LogTabsState(L"OnTabControlCloseRequested");
}

void TabBarHost::OnTabControlContextMenuRequested(std::optional<std::size_t> /*index*/, const POINT& screenPoint) {
    ATLTRACE(L"TabBarHost::OnTabControlContextMenuRequested (%ld,%ld)\n", screenPoint.x, screenPoint.y);
    ShowContextMenu(screenPoint);
}

void TabBarHost::OnTabControlNewTabRequested() {
    ATLTRACE(L"TabBarHost::OnTabControlNewTabRequested\n");
    AddTab(m_currentPath, true, true);
}

void TabBarHost::OnTabControlBeginDrag(std::size_t index, const POINT& screenPoint) {
    ATLTRACE(L"TabBarHost::OnTabControlBeginDrag index=%u (%ld,%ld)\n", static_cast<UINT>(index), screenPoint.x, screenPoint.y);
}
