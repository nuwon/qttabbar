#include "pch.h"
#include "QTButtonBar.h"

#include <algorithm>
#include <array>
#include <cwchar>

#include "InstanceManagerNative.h"
#include "QTTabBarClass.h"

using qttabbar::InstanceManagerNative;

namespace {

constexpr UINT kToolbarPadding = 4;
constexpr UINT kSearchBoxMinWidth = 120;

const std::array<QTButtonBar::ButtonDefinition, 21> kButtonDefinitions = {{
    {QTButtonBar::BII_NAVIGATION_BACK, QTButtonBar::CMD_NAV_BACK, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_BACK},
    {QTButtonBar::BII_NAVIGATION_FORWARD, QTButtonBar::CMD_NAV_FORWARD, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_FORWARD},
    {QTButtonBar::BII_GROUP, QTButtonBar::CMD_GROUPS, BTNS_DROPDOWN | BTNS_AUTOSIZE, IDS_BUTTONBAR_GROUPS},
    {QTButtonBar::BII_RECENTTAB, QTButtonBar::CMD_RECENT_TABS, BTNS_DROPDOWN | BTNS_AUTOSIZE, IDS_BUTTONBAR_RECENT_TABS},
    {QTButtonBar::BII_APPLICATIONLAUNCHER, QTButtonBar::CMD_APP_LAUNCHER, BTNS_DROPDOWN | BTNS_AUTOSIZE, IDS_BUTTONBAR_APPS},
    {QTButtonBar::BII_NEWWINDOW, QTButtonBar::CMD_NEW_WINDOW, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_NEW_WINDOW},
    {QTButtonBar::BII_CLONE, QTButtonBar::CMD_CLONE_TAB, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_CLONE},
    {QTButtonBar::BII_LOCK, QTButtonBar::CMD_LOCK_TAB, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_LOCK},
    {QTButtonBar::BII_MISCTOOL, QTButtonBar::CMD_MISC_TOOLS, BTNS_DROPDOWN | BTNS_AUTOSIZE, IDS_BUTTONBAR_MISC},
    {QTButtonBar::BII_TOPMOST, QTButtonBar::CMD_TOGGLE_TOPMOST, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_TOPMOST},
    {QTButtonBar::BII_CLOSE_CURRENT, QTButtonBar::CMD_CLOSE_CURRENT, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_CLOSE},
    {QTButtonBar::BII_CLOSE_ALLBUTCURRENT, QTButtonBar::CMD_CLOSE_OTHERS, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_CLOSE_OTHERS},
    {QTButtonBar::BII_CLOSE_WINDOW, QTButtonBar::CMD_CLOSE_WINDOW, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_CLOSE_WINDOW},
    {QTButtonBar::BII_CLOSE_LEFT, QTButtonBar::CMD_CLOSE_LEFT, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_CLOSE_LEFT},
    {QTButtonBar::BII_CLOSE_RIGHT, QTButtonBar::CMD_CLOSE_RIGHT, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_CLOSE_RIGHT},
    {QTButtonBar::BII_GOUPONELEVEL, QTButtonBar::CMD_GO_UP, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_GO_UP},
    {QTButtonBar::BII_REFRESH_SHELLBROWSER, QTButtonBar::CMD_REFRESH, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_REFRESH},
    {QTButtonBar::BII_SHELLSEARCH, QTButtonBar::CMD_SEARCH, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_SEARCH},
    {QTButtonBar::BII_WINDOWOPACITY, QTButtonBar::CMD_WINDOW_OPACITY, BTNS_DROPDOWN | BTNS_AUTOSIZE, IDS_BUTTONBAR_OPACITY},
    {QTButtonBar::BII_FILTERBAR, QTButtonBar::CMD_FILTER_BAR, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_FILTER},
    {QTButtonBar::BII_OPTION, QTButtonBar::CMD_OPTIONS, BTNS_BUTTON | BTNS_AUTOSIZE, IDS_BUTTONBAR_OPTIONS},
}};

const QTButtonBar::ButtonDefinition* FindButtonDefinition(QTButtonBar::ButtonIndex index) {
    for(const auto& def : kButtonDefinitions) {
        if(def.index == index) {
            return &def;
        }
    }
    return nullptr;
}

std::wstring LoadStringResource(UINT id) {
    ATL::CStringW value;
    value.LoadStringW(id);
    return std::wstring(value.GetString(), value.GetString() + value.GetLength());
}

}  // namespace

QTButtonBar::QTButtonBar() noexcept
    : m_hwndToolbar(nullptr)
    , m_hwndSearch(nullptr)
    , m_hImageList(nullptr)
    , m_minSize{160, 28}
    , m_maxSize{-1, -1}
    , m_visible(false)
    , m_hasFocus(false) {}

QTButtonBar::~QTButtonBar() = default;

HRESULT QTButtonBar::FinalConstruct() {
    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_BAR_CLASSES | ICC_COOL_CLASSES;
    if(!::InitCommonControlsEx(&icc)) {
        return HRESULT_FROM_WIN32(::GetLastError());
    }
    m_config = qttabbar::LoadConfigFromRegistry();
    return S_OK;
}

void QTButtonBar::FinalRelease() {
    DestroyWindowResources();
    m_spExplorer.Release();
    m_spServiceProvider.Release();
    m_spInputObjectSite.Release();
    m_spSite.Release();
}

IFACEMETHODIMP QTButtonBar::GetWindow(HWND* phwnd) {
    if(phwnd == nullptr) {
        return E_POINTER;
    }
    *phwnd = m_hWnd;
    return (m_hWnd != nullptr) ? S_OK : E_FAIL;
}

IFACEMETHODIMP QTButtonBar::ContextSensitiveHelp(BOOL) {
    return E_NOTIMPL;
}

IFACEMETHODIMP QTButtonBar::ShowDW(BOOL fShow) {
    if(fShow && m_hWnd == nullptr) {
        EnsureWindow();
    }
    m_visible = fShow != FALSE;
    if(m_hWnd != nullptr) {
        ::ShowWindow(m_hWnd, fShow ? SW_SHOW : SW_HIDE);
    }
    if(m_hwndToolbar != nullptr) {
        ::ShowWindow(m_hwndToolbar, fShow ? SW_SHOW : SW_HIDE);
    }
    if(m_hwndSearch != nullptr) {
        ::ShowWindow(m_hwndSearch, fShow ? SW_SHOW : SW_HIDE);
    }
    return S_OK;
}

IFACEMETHODIMP QTButtonBar::CloseDW(DWORD) {
    DestroyWindowResources();
    return S_OK;
}

IFACEMETHODIMP QTButtonBar::ResizeBorderDW(LPCRECT, IUnknown*, BOOL) {
    return E_NOTIMPL;
}

IFACEMETHODIMP QTButtonBar::GetBandInfo(DWORD dwBandID, DWORD, DESKBANDINFO* pdbi) {
    if(pdbi == nullptr) {
        return E_POINTER;
    }
    pdbi->dwMask &= ~(DBIM_ACTUAL | DBIM_BKCOLOR | DBIM_MODEFLAGS | DBIM_MINMAX);
    pdbi->dwMask |= DBIM_ACTUAL | DBIM_MODEFLAGS | DBIM_MINMAX;
    pdbi->dwBandID = dwBandID;
    pdbi->ptMinSize.x = m_minSize.cx;
    pdbi->ptMinSize.y = m_minSize.cy;
    pdbi->ptIntegral.x = 1;
    pdbi->ptIntegral.y = 1;
    pdbi->ptActual.x = m_minSize.cx;
    pdbi->ptActual.y = m_minSize.cy;
    pdbi->ptMaxSize.x = m_maxSize.cx;
    pdbi->ptMaxSize.y = m_maxSize.cy;
    pdbi->dwModeFlags = DBIMF_NORMAL | DBIMF_USECHEVRON;
    return S_OK;
}

IFACEMETHODIMP QTButtonBar::UIActivateIO(BOOL fActivate, MSG* pMsg) {
    if(fActivate && m_hWnd != nullptr) {
        ::SetFocus(m_hWnd);
        if(pMsg != nullptr) {
            TranslateMessage(pMsg);
            DispatchMessage(pMsg);
        }
        return S_OK;
    }
    return S_FALSE;
}

IFACEMETHODIMP QTButtonBar::HasFocusIO() {
    return m_hasFocus ? S_OK : S_FALSE;
}

IFACEMETHODIMP QTButtonBar::TranslateAcceleratorIO(MSG* pMsg) {
    if(pMsg == nullptr) {
        return E_POINTER;
    }
    if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_RETURN && m_hwndSearch != nullptr) {
        if(::GetFocus() == m_hwndSearch) {
            HandleCommand(CMD_SEARCH);
            return S_OK;
        }
    }
    return S_FALSE;
}

IFACEMETHODIMP QTButtonBar::SetSite(IUnknown* pUnkSite) {
    if(pUnkSite == nullptr) {
        InstanceManagerNative::UnregisterButtonBar(this);
        m_spInputObjectSite.Release();
        m_spServiceProvider.Release();
        m_spExplorer.Release();
        m_spSite.Release();
        return S_OK;
    }

    m_spSite = pUnkSite;
    m_spInputObjectSite.Release();
    m_spServiceProvider.Release();
    m_spExplorer.Release();

    pUnkSite->QueryInterface(IID_PPV_ARGS(&m_spInputObjectSite));
    pUnkSite->QueryInterface(IID_PPV_ARGS(&m_spServiceProvider));

    if(m_spServiceProvider) {
        m_spServiceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&m_spExplorer));
    }

    InstanceManagerNative::RegisterButtonBar(this);
    return S_OK;
}

IFACEMETHODIMP QTButtonBar::GetSite(REFIID riid, void** ppvSite) {
    if(ppvSite == nullptr) {
        return E_POINTER;
    }
    if(!m_spSite) {
        *ppvSite = nullptr;
        return E_FAIL;
    }
    return m_spSite->QueryInterface(riid, ppvSite);
}

IFACEMETHODIMP QTButtonBar::GetClassID(CLSID* pClassID) {
    if(pClassID == nullptr) {
        return E_POINTER;
    }
    *pClassID = CLSID_QTButtonBar;
    return S_OK;
}

IFACEMETHODIMP QTButtonBar::IsDirty() {
    return S_FALSE;
}

IFACEMETHODIMP QTButtonBar::Load(IStream*) {
    return S_OK;
}

IFACEMETHODIMP QTButtonBar::Save(IStream*, BOOL) {
    return S_OK;
}

IFACEMETHODIMP QTButtonBar::GetSizeMax(ULARGE_INTEGER* pcbSize) {
    if(pcbSize == nullptr) {
        return E_POINTER;
    }
    pcbSize->QuadPart = 0;
    return E_NOTIMPL;
}

LRESULT QTButtonBar::OnCreate(UINT, WPARAM, LPARAM, BOOL& bHandled) {
    bHandled = TRUE;
    InitializeToolbar();
    LayoutControls();
    return 0;
}

LRESULT QTButtonBar::OnDestroy(UINT, WPARAM, LPARAM, BOOL& bHandled) {
    bHandled = TRUE;
    InstanceManagerNative::UnregisterButtonBar(this);
    DestroyWindowResources();
    return 0;
}

LRESULT QTButtonBar::OnSize(UINT, WPARAM, LPARAM, BOOL& bHandled) {
    bHandled = TRUE;
    LayoutControls();
    return 0;
}

LRESULT QTButtonBar::OnCommand(UINT, WPARAM wParam, LPARAM, BOOL& bHandled) {
    bHandled = TRUE;
    HandleCommand(LOWORD(wParam));
    return 0;
}

LRESULT QTButtonBar::OnNotify(UINT, WPARAM, LPARAM lParam, BOOL& bHandled) {
    auto* hdr = reinterpret_cast<NMHDR*>(lParam);
    if(hdr == nullptr) {
        bHandled = FALSE;
        return 0;
    }
    if(hdr->hwndFrom == m_hwndToolbar && hdr->code == TBN_DROPDOWN) {
        auto* info = reinterpret_cast<NMTOOLBARW*>(lParam);
        RECT rc{};
        ::SendMessageW(m_hwndToolbar, TB_GETRECT, info->iItem, reinterpret_cast<LPARAM>(&rc));
        ::MapWindowPoints(m_hwndToolbar, HWND_DESKTOP, reinterpret_cast<POINT*>(&rc), 2);
        ShowDropdownMenu(static_cast<UINT>(info->iItem), rc);
        bHandled = TRUE;
        return TBDDRET_DEFAULT;
    }
    bHandled = FALSE;
    return 0;
}

LRESULT QTButtonBar::OnSetFocus(UINT, WPARAM, LPARAM, BOOL& bHandled) {
    bHandled = TRUE;
    m_hasFocus = true;
    if(m_spInputObjectSite) {
        m_spInputObjectSite->OnFocusChangeIS(this, TRUE);
    }
    return 0;
}

LRESULT QTButtonBar::OnKillFocus(UINT, WPARAM, LPARAM, BOOL& bHandled) {
    bHandled = TRUE;
    m_hasFocus = false;
    if(m_spInputObjectSite) {
        m_spInputObjectSite->OnFocusChangeIS(this, FALSE);
    }
    return 0;
}

HRESULT QTButtonBar::EnsureWindow() {
    if(m_hWnd != nullptr) {
        return S_OK;
    }
    HWND parent = nullptr;
    if(m_spSite) {
        CComPtr<IOleWindow> spOleWindow;
        if(SUCCEEDED(m_spSite->QueryInterface(IID_PPV_ARGS(&spOleWindow))) && spOleWindow) {
            spOleWindow->GetWindow(&parent);
        }
    }
    if(parent == nullptr) {
        return E_UNEXPECTED;
    }
    RECT rc{0, 0, m_minSize.cx, m_minSize.cy};
    HWND hwnd = Create(parent, rc, L"", WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS);
    if(hwnd == nullptr) {
        return HRESULT_FROM_WIN32(::GetLastError());
    }
    return S_OK;
}

void QTButtonBar::DestroyWindowResources() {
    if(m_hwndToolbar != nullptr) {
        ::DestroyWindow(m_hwndToolbar);
        m_hwndToolbar = nullptr;
    }
    if(m_hwndSearch != nullptr) {
        ::DestroyWindow(m_hwndSearch);
        m_hwndSearch = nullptr;
    }
    if(m_hImageList != nullptr) {
        ::ImageList_Destroy(m_hImageList);
        m_hImageList = nullptr;
    }
    if(m_hWnd != nullptr) {
        ::DestroyWindow(m_hWnd);
        m_hWnd = nullptr;
    }
}

void QTButtonBar::InitializeToolbar() {
    if(m_hWnd == nullptr) {
        EnsureWindow();
    }
    if(m_hWnd == nullptr) {
        return;
    }
    if(m_hwndToolbar == nullptr) {
        m_hwndToolbar = ::CreateWindowExW(0, TOOLBARCLASSNAMEW, L"",
                                          WS_CHILD | WS_VISIBLE | TBSTYLE_FLAT | TBSTYLE_LIST | TBSTYLE_TOOLTIPS |
                                              CCS_NOPARENTALIGN | CCS_NORESIZE | CCS_NODIVIDER,
                                          0, 0, 0, 0, m_hWnd, nullptr, _AtlBaseModule.GetModuleInstance(), nullptr);
        if(m_hwndToolbar != nullptr) {
            ::SendMessageW(m_hwndToolbar, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);
            ::SendMessageW(m_hwndToolbar, TB_SETEXTENDEDSTYLE, 0, TBSTYLE_EX_DOUBLEBUFFER | TBSTYLE_EX_MIXEDBUTTONS);
        }
    }
    if(m_hwndSearch == nullptr) {
        m_hwndSearch = ::CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                         0, 0, 0, 0, m_hWnd, reinterpret_cast<HMENU>(IDC_BUTTONBAR_SEARCHBOX),
                                         _AtlBaseModule.GetModuleInstance(), nullptr);
    }
    PopulateButtons();
}

void QTButtonBar::PopulateButtons() {
    if(m_hwndToolbar == nullptr) {
        return;
    }
    int count = static_cast<int>(::SendMessageW(m_hwndToolbar, TB_BUTTONCOUNT, 0, 0));
    for(int i = count - 1; i >= 0; --i) {
        ::SendMessageW(m_hwndToolbar, TB_DELETEBUTTON, i, 0);
    }
    for(int indexValue : m_config.bbar.buttonIndexes) {
        ButtonIndex index = static_cast<ButtonIndex>(indexValue);
        if(index == BII_SEPARATOR) {
            AddSeparator();
            continue;
        }
        const ButtonDefinition* def = FindButtonDefinition(index);
        if(def == nullptr) {
            continue;
        }
        AddButton(*def);
    }
    ::SendMessageW(m_hwndToolbar, TB_AUTOSIZE, 0, 0);
}

void QTButtonBar::AddButton(const ButtonDefinition& def) {
    if(m_hwndToolbar == nullptr) {
        return;
    }
    std::wstring label = LoadStringResource(def.stringResId);
    LRESULT stringIndex = ::SendMessageW(m_hwndToolbar, TB_ADDSTRINGW, 0, reinterpret_cast<LPARAM>(label.c_str()));
    TBBUTTON button{};
    button.iBitmap = I_IMAGENONE;
    button.idCommand = def.command;
    button.fsState = TBSTATE_ENABLED;
    button.fsStyle = static_cast<BYTE>(def.style);
    button.dwData = static_cast<DWORD_PTR>(def.index);
    button.iString = static_cast<int>(stringIndex);
    ::SendMessageW(m_hwndToolbar, TB_ADDBUTTONSW, 1, reinterpret_cast<LPARAM>(&button));
}

void QTButtonBar::AddSeparator() {
    if(m_hwndToolbar == nullptr) {
        return;
    }
    TBBUTTON button{};
    button.fsStyle = BTNS_SEP;
    ::SendMessageW(m_hwndToolbar, TB_ADDBUTTONSW, 1, reinterpret_cast<LPARAM>(&button));
}

void QTButtonBar::LayoutControls() {
    if(m_hWnd == nullptr) {
        return;
    }
    RECT rc{};
    ::GetClientRect(m_hWnd, &rc);
    int width = rc.right - rc.left;
    int height = rc.bottom - rc.top;
    int toolbarWidth = 0;
    if(m_hwndToolbar != nullptr) {
        SIZE size{};
        if(::SendMessageW(m_hwndToolbar, TB_GETMAXSIZE, 0, reinterpret_cast<LPARAM>(&size))) {
            toolbarWidth = static_cast<int>(size.cx);
        }
        ::MoveWindow(m_hwndToolbar, kToolbarPadding, 0, toolbarWidth, height, TRUE);
    }
    int searchX = toolbarWidth + (kToolbarPadding * 2);
    int searchWidth = std::max<int>(kSearchBoxMinWidth, width - searchX - kToolbarPadding);
    if(m_hwndSearch != nullptr) {
        ::MoveWindow(m_hwndSearch, searchX, 2, searchWidth, height - 4, TRUE);
    }
}

void QTButtonBar::HandleCommand(UINT commandId) {
    switch(commandId) {
    case CMD_NAV_BACK:
        if(m_spExplorer) {
            m_spExplorer->GoBack();
        }
        break;
    case CMD_NAV_FORWARD:
        if(m_spExplorer) {
            m_spExplorer->GoForward();
        }
        break;
    case CMD_GROUPS:
    case CMD_RECENT_TABS:
    case CMD_APP_LAUNCHER:
    case CMD_MISC_TOOLS:
    case CMD_WINDOW_OPACITY:
        if(m_hwndToolbar) {
            RECT rc{};
            ::SendMessageW(m_hwndToolbar, TB_GETRECT, commandId, reinterpret_cast<LPARAM>(&rc));
            ::MapWindowPoints(m_hwndToolbar, HWND_DESKTOP, reinterpret_cast<POINT*>(&rc), 2);
            ShowDropdownMenu(commandId, rc);
        }
        break;
    case CMD_NEW_WINDOW: {
        CComQIPtr<IOleCommandTarget> commandTarget(m_spExplorer);
        if(commandTarget) {
            commandTarget->Exec(&CGID_ShellDocView, OLECMDID_NEW, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, nullptr);
        }
        break;
    }
    case CMD_CLONE_TAB:
        DispatchToTabHost(ID_CONTEXT_NEWTAB);
        break;
    case CMD_LOCK_TAB:
        // Placeholder for future lock behavior.
        break;
    case CMD_TOGGLE_TOPMOST:
        // Placeholder for toggling topmost state.
        break;
    case CMD_CLOSE_CURRENT:
        DispatchToTabHost(ID_CONTEXT_CLOSETAB);
        break;
    case CMD_CLOSE_OTHERS:
    case CMD_CLOSE_LEFT:
    case CMD_CLOSE_RIGHT:
        // Future enhancements will route these to the tab host.
        break;
    case CMD_CLOSE_WINDOW:
        if(m_spExplorer) {
            m_spExplorer->Quit();
        }
        break;
    case CMD_GO_UP:
        if(m_spExplorer) {
            CComQIPtr<IShellBrowser> shellBrowser(m_spExplorer);
            if(shellBrowser) {
                shellBrowser->BrowseObject(nullptr, SBSP_PARENT);
            }
        }
        break;
    case CMD_REFRESH:
        DispatchToTabHost(ID_CONTEXT_REFRESH);
        break;
    case CMD_SEARCH:
        if(m_spExplorer && m_hwndSearch) {
            int length = ::GetWindowTextLengthW(m_hwndSearch);
            if(length > 0) {
                std::wstring buffer(static_cast<size_t>(length) + 1, L'\0');
                ::GetWindowTextW(m_hwndSearch, buffer.data(), length + 1);
                buffer.resize(wcslen(buffer.c_str()));
                if(!buffer.empty()) {
                    std::wstring searchUrl = L"search-ms:?query=" + buffer;
                    CComVariant url(searchUrl.c_str());
                    m_spExplorer->Navigate2(&url, nullptr, nullptr, nullptr, nullptr);
                }
            }
        }
        break;
    case CMD_FILTER_BAR:
        // Placeholder for filter bar toggle.
        break;
    case CMD_OPTIONS:
        DispatchToTabHost(ID_CONTEXT_OPTIONS);
        break;
    default:
        break;
    }
}

void QTButtonBar::ShowDropdownMenu(UINT commandId, const RECT& buttonRect) {
    HMENU menu = ::CreatePopupMenu();
    if(menu == nullptr) {
        return;
    }
    auto appendPlaceholder = [menu](UINT stringId) {
        std::wstring text = LoadStringResource(stringId);
        ::AppendMenuW(menu, MF_STRING | MF_GRAYED, 0, text.c_str());
    };

    switch(commandId) {
    case CMD_GROUPS:
        appendPlaceholder(IDS_BUTTONBAR_GROUPS_PLACEHOLDER);
        break;
    case CMD_RECENT_TABS:
        appendPlaceholder(IDS_BUTTONBAR_RECENT_PLACEHOLDER);
        break;
    case CMD_APP_LAUNCHER:
        if(m_config.bbar.activePluginIDs.empty()) {
            appendPlaceholder(IDS_BUTTONBAR_NO_PLUGINS);
        } else {
            UINT id = CMD_APP_LAUNCHER + 0x100;
            for(const auto& pluginId : m_config.bbar.activePluginIDs) {
                ::AppendMenuW(menu, MF_STRING, id++, pluginId.c_str());
            }
        }
        break;
    case CMD_MISC_TOOLS:
        appendPlaceholder(IDS_BUTTONBAR_MISC_PLACEHOLDER);
        break;
    case CMD_WINDOW_OPACITY:
        appendPlaceholder(IDS_BUTTONBAR_OPACITY_PLACEHOLDER);
        break;
    default:
        break;
    }

    ::TrackPopupMenuEx(menu, TPM_LEFTALIGN | TPM_LEFTBUTTON, buttonRect.left, buttonRect.bottom, m_hWnd, nullptr);
    ::DestroyMenu(menu);
}

void QTButtonBar::DispatchToTabHost(UINT commandId) {
    QTTabBarClass* tabBar = InstanceManagerNative::GetActiveTabBar();
    if(tabBar != nullptr) {
        tabBar->ExecuteHostCommand(commandId);
    }
}

