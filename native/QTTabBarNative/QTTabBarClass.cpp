#include "pch.h"
#include "QTTabBarClass.h"
#include "NativeTabControl.h"
#include "NativeTabNotifications.h"

namespace {

constexpr DWORD kRebarMaskStyle = RBBIM_STYLE | RBBIM_CHILD;

class RebarBreakFixer {
public:
    RebarBreakFixer(HWND hwndRebar, QTTabBarClass* parent)
        : m_hwnd(hwndRebar)
        , m_parent(parent)
        , m_prevProc(nullptr)
        , m_monitorSetInfo(true)
        , m_enabled(true) {
        ATLASSERT(::IsWindow(hwndRebar));
        std::scoped_lock lock(s_mapMutex);
        m_prevProc = reinterpret_cast<WNDPROC>(::SetWindowLongPtrW(m_hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(&RebarBreakFixer::SubclassProc)));
        s_instances[m_hwnd] = this;
    }

    ~RebarBreakFixer() {
        std::scoped_lock lock(s_mapMutex);
        if(::IsWindow(m_hwnd) && m_prevProc != nullptr) {
            ::SetWindowLongPtrW(m_hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(m_prevProc));
        }
        s_instances.erase(m_hwnd);
    }

    RebarBreakFixer(const RebarBreakFixer&) = delete;
    RebarBreakFixer& operator=(const RebarBreakFixer&) = delete;

    void SetMonitorSetInfo(bool enabled) noexcept { m_monitorSetInfo = enabled; }
    void SetEnabled(bool enabled) noexcept { m_enabled = enabled; }

private:
    static RebarBreakFixer* Lookup(HWND hwnd) {
        std::scoped_lock lock(s_mapMutex);
        auto it = s_instances.find(hwnd);
        return it != s_instances.end() ? it->second : nullptr;
    }

    static LRESULT CALLBACK SubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
        RebarBreakFixer* self = Lookup(hwnd);
        if(self == nullptr) {
            return ::DefWindowProcW(hwnd, msg, wParam, lParam);
        }

        if(!self->m_enabled) {
            return ::CallWindowProcW(self->m_prevProc, hwnd, msg, wParam, lParam);
        }

        if(msg == RB_SETBANDINFO && self->m_monitorSetInfo && lParam != 0) {
            auto* bandInfo = reinterpret_cast<REBARBANDINFOW*>(lParam);
            if((bandInfo->fMask & RBBIM_STYLE) != 0 && bandInfo->hwndChild == self->m_parent->m_hWnd) {
                if(self->m_parent->ShouldHaveBreak()) {
                    bandInfo->fStyle |= RBBS_BREAK;
                } else {
                    bandInfo->fStyle &= ~RBBS_BREAK;
                }
            }
        } else if(msg == RB_DELETEBAND) {
            int deleteIndex = static_cast<int>(wParam);
            int count = self->m_parent->ActiveRebarCount();
            for(int i = 0; i < count; ++i) {
                REBARBANDINFOW info = self->m_parent->GetRebarBand(i, kRebarMaskStyle | RBBIM_STYLE);
                if(info.hwndChild == self->m_parent->m_hWnd) {
                    LRESULT result = ::CallWindowProcW(self->m_prevProc, hwnd, msg, wParam, lParam);
                    if(i == deleteIndex) {
                        return result;
                    }

                    REBARBANDINFOW restore = {};
                    restore.cbSize = sizeof(restore);
                    restore.fMask = RBBIM_STYLE;
                    restore.fStyle = info.fStyle;

                    bool previousMonitor = self->m_monitorSetInfo;
                    self->m_monitorSetInfo = false;
                    ::SendMessageW(hwnd, RB_SETBANDINFO, static_cast<WPARAM>(i), reinterpret_cast<LPARAM>(&restore));
                    self->m_monitorSetInfo = previousMonitor;
                    return result;
                }
            }
        }

        return ::CallWindowProcW(self->m_prevProc, hwnd, msg, wParam, lParam);
    }

    HWND m_hwnd;
    QTTabBarClass* m_parent;
    WNDPROC m_prevProc;
    bool m_monitorSetInfo;
    bool m_enabled;

    static std::mutex s_mapMutex;
    static std::unordered_map<HWND, RebarBreakFixer*> s_instances;
};

std::mutex RebarBreakFixer::s_mapMutex;
std::unordered_map<HWND, RebarBreakFixer*> RebarBreakFixer::s_instances;

} // namespace

QTTabBarClass::QTTabBarClass() noexcept
    : m_hwndRebar(nullptr)
    , m_hwndHost(nullptr)
    , m_hContextMenu(nullptr)
    , m_minSize{16, 26}
    , m_maxSize{-1, -1}
    , m_closed(false)
    , m_visible(false)
    , m_vertical(false)
    , m_bandId(0) {
}

QTTabBarClass::~QTTabBarClass() {
}

HRESULT QTTabBarClass::FinalConstruct() {
    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_BAR_CLASSES | ICC_TAB_CLASSES | ICC_LISTVIEW_CLASSES;
    ::InitCommonControlsEx(&icc);
    return S_OK;
}

void QTTabBarClass::FinalRelease() {
    DestroyTimers();
    DestroyContextMenu();
    ReleaseRebarSubclass();
    if(m_tabControl) {
        if(m_tabControl->m_hWnd != nullptr && ::IsWindow(m_tabControl->m_hWnd)) {
            ::DestroyWindow(m_tabControl->m_hWnd);
        }
        m_tabControl.reset();
    }
    m_hwndHost = nullptr;
    if(m_hWnd != nullptr && ::IsWindow(m_hWnd)) {
        ::DestroyWindow(m_hWnd);
        m_hWnd = nullptr;
    }
    m_spExplorer.Release();
    m_spServiceProvider.Release();
    m_spInputObjectSite.Release();
    m_spSite.Release();
}

HRESULT QTTabBarClass::EnsureWindow() {
    if(m_hWnd != nullptr) {
        return S_OK;
    }

    if(m_hwndRebar == nullptr) {
        return E_UNEXPECTED;
    }

    RECT rc = {0, 0, m_minSize.cx, m_minSize.cy};
    HWND hwnd = Create(m_hwndRebar, rc, L"", WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS);
    if(hwnd == nullptr) {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    return S_OK;
}

void QTTabBarClass::InitializeTimers() {
    if(m_hWnd == nullptr) {
        return;
    }
    ::SetTimer(m_hWnd, ID_TIMER_SELECTTAB, kSelectTabTimerMs, nullptr);
    ::SetTimer(m_hWnd, ID_TIMER_CONTEXTMENU, kShowMenuTimerMs, nullptr);
}

void QTTabBarClass::DestroyTimers() {
    if(m_hWnd != nullptr) {
        ::KillTimer(m_hWnd, ID_TIMER_SELECTTAB);
        ::KillTimer(m_hWnd, ID_TIMER_CONTEXTMENU);
    }
}

void QTTabBarClass::InitializeContextMenu() {
    if(m_hContextMenu != nullptr) {
        return;
    }

    HINSTANCE module = _AtlBaseModule.GetModuleInstance();
    HMENU menu = ::LoadMenuW(module, MAKEINTRESOURCEW(IDR_QTTABBAR_CONTEXT_MENU));
    if(menu != nullptr) {
        m_hContextMenu = menu;
    }
}

void QTTabBarClass::DestroyContextMenu() {
    if(m_hContextMenu != nullptr) {
        ::DestroyMenu(m_hContextMenu);
        m_hContextMenu = nullptr;
    }
}

void QTTabBarClass::EnsureRebarSubclass() {
    if(m_hwndRebar == nullptr) {
        return;
    }
    if(!m_rebarSubclass) {
        m_rebarSubclass = std::make_unique<RebarBreakFixer>(m_hwndRebar, this);
    }
    if(m_rebarSubclass) {
        m_rebarSubclass->SetEnabled(true);
        m_rebarSubclass->SetMonitorSetInfo(true);
    }
}

void QTTabBarClass::ReleaseRebarSubclass() {
    if(m_rebarSubclass) {
        m_rebarSubclass->SetEnabled(false);
        m_rebarSubclass.reset();
    }
}

void QTTabBarClass::UpdateVisibility(BOOL fShow) {
    m_visible = fShow != FALSE;
    if(m_hWnd != nullptr) {
        ::ShowWindow(m_hWnd, fShow ? SW_SHOW : SW_HIDE);
        if(fShow) {
            ::SetWindowPos(m_hWnd, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
    }
    if(fShow) {
        EnsureRebarSubclass();
        StartDeferredRebarReset();
    }
}

void QTTabBarClass::NotifyFocusChange(BOOL hasFocus) {
    if(m_spInputObjectSite != nullptr && !m_closed) {
        m_spInputObjectSite->OnFocusChangeIS(this, hasFocus);
    }
}

void QTTabBarClass::HandleContextCommand(UINT commandId, UINT tabIndex) {
    ATLTRACE(L"QTTabBarClass::HandleContextCommand %u\n", commandId);
    switch(commandId) {
    case ID_CONTEXT_NEWTAB:
    case ID_CONTEXT_CLOSETAB:
    case ID_CONTEXT_OPTIONS:
        if(m_hwndRebar != nullptr) {
            ::SendMessageW(m_hwndRebar, WM_COMMAND, MAKEWPARAM(commandId, tabIndex), reinterpret_cast<LPARAM>(m_hWnd));
        }
        break;
    case ID_CONTEXT_REFRESH:
        if(m_spExplorer) {
            m_spExplorer->Refresh();
        }
        break;
    default:
        break;
    }
}

void QTTabBarClass::StartDeferredRebarReset() {
    if(m_hWnd != nullptr) {
        ::PostMessageW(m_hWnd, WM_APP_UNSUBCLASS, 0, 0);
    }
}

bool QTTabBarClass::ShouldHaveBreak() const {
    return true;
}

int QTTabBarClass::ActiveRebarCount() const {
    if(m_hwndRebar == nullptr) {
        return 0;
    }
    return static_cast<int>(::SendMessageW(m_hwndRebar, RB_GETBANDCOUNT, 0, 0));
}

REBARBANDINFOW QTTabBarClass::GetRebarBand(int index, UINT mask) const {
    REBARBANDINFOW info = {};
    if(m_hwndRebar != nullptr) {
        info.cbSize = sizeof(info);
        info.fMask = mask;
        ::SendMessageW(m_hwndRebar, RB_GETBANDINFO, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(&info));
    }
    return info;
}

bool QTTabBarClass::BandHasBreak() const {
    int count = ActiveRebarCount();
    for(int i = 0; i < count; ++i) {
        REBARBANDINFOW info = GetRebarBand(i, kRebarMaskStyle | RBBIM_STYLE);
        if(info.hwndChild == m_hWnd) {
            return (info.fStyle & RBBS_BREAK) != 0;
        }
    }
    return true;
}

LRESULT QTTabBarClass::OnCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;

    RECT rc{0, 0, m_minSize.cx, m_minSize.cy};
    m_tabControl = std::make_unique<NativeTabControl>();
    if(m_tabControl) {
        HWND hwndTab = m_tabControl->Create(m_hWnd, rc, nullptr,
                                           WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_CLIPSIBLINGS | WS_CLIPCHILDREN,
                                           0, ID_TAB_CONTROL);
        m_hwndHost = hwndTab;
    }
    InitializeContextMenu();
    InitializeTimers();
    return 0;
}

LRESULT QTTabBarClass::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    DestroyTimers();
    DestroyContextMenu();
    if(m_tabControl) {
        if(m_tabControl->m_hWnd != nullptr && ::IsWindow(m_tabControl->m_hWnd)) {
            ::DestroyWindow(m_tabControl->m_hWnd);
        }
        m_tabControl.reset();
    }
    m_hwndHost = nullptr;
    return 0;
}

LRESULT QTTabBarClass::OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    if(m_tabControl && m_tabControl->m_hWnd != nullptr) {
        int width = LOWORD(lParam);
        int height = HIWORD(lParam);
        ::SetWindowPos(m_tabControl->m_hWnd, nullptr, 0, 0, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
        m_tabControl->EnsureLayout();
    }
    return 0;
}

LRESULT QTTabBarClass::OnTimer(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    if(wParam == ID_TIMER_SELECTTAB) {
        EnsureRebarSubclass();
    } else if(wParam == ID_TIMER_CONTEXTMENU) {
        StartDeferredRebarReset();
    }
    return 0;
}

LRESULT QTTabBarClass::OnContextMenu(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    if(m_hContextMenu == nullptr) {
        return 0;
    }

    HMENU subMenu = ::GetSubMenu(m_hContextMenu, 0);
    if(subMenu == nullptr) {
        return 0;
    }

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

    ::TrackPopupMenuEx(subMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON, pt.x, pt.y, m_hWnd, nullptr);
    return 0;
}

LRESULT QTTabBarClass::OnCommand(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    UINT commandId = LOWORD(wParam);
    UINT tabIndex = HIWORD(wParam);
    HWND source = reinterpret_cast<HWND>(lParam);
    if(source == m_hwndHost && source != nullptr) {
        HandleContextCommand(commandId, tabIndex);
        return 0;
    }
    HandleContextCommand(commandId, tabIndex);
    return 0;
}

LRESULT QTTabBarClass::OnNotify(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = FALSE;
    NMHDR* hdr = reinterpret_cast<NMHDR*>(lParam);
    if(hdr != nullptr && m_tabControl) {
        HWND tooltip = m_tabControl->GetTooltipWindow();
        if(hdr->hwndFrom == m_tabControl->m_hWnd || (tooltip != nullptr && hdr->hwndFrom == tooltip)) {
            if(m_hwndRebar != nullptr) {
                ::SendMessageW(m_hwndRebar, WM_NOTIFY, wParam, lParam);
            }
            bHandled = TRUE;
            return 0;
        }
    }
    return 0;
}

LRESULT QTTabBarClass::OnSetFocus(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    NotifyFocusChange(TRUE);
    return 0;
}

LRESULT QTTabBarClass::OnKillFocus(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    NotifyFocusChange(FALSE);
    return 0;
}

LRESULT QTTabBarClass::OnUnsetRebarMonitor(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    if(m_rebarSubclass) {
        m_rebarSubclass->SetMonitorSetInfo(false);
    }
    return 0;
}

IFACEMETHODIMP QTTabBarClass::GetWindow(HWND* phwnd) {
    if(phwnd == nullptr) {
        return E_POINTER;
    }
    HRESULT hr = EnsureWindow();
    if(FAILED(hr)) {
        return hr;
    }
    *phwnd = m_hWnd;
    return S_OK;
}

IFACEMETHODIMP QTTabBarClass::ContextSensitiveHelp(BOOL /*fEnterMode*/) {
    return E_NOTIMPL;
}

IFACEMETHODIMP QTTabBarClass::ShowDW(BOOL fShow) {
    HRESULT hr = EnsureWindow();
    if(FAILED(hr)) {
        return hr;
    }
    UpdateVisibility(fShow);
    return S_OK;
}

IFACEMETHODIMP QTTabBarClass::CloseDW(DWORD /*dwReserved*/) {
    m_closed = true;
    ShowDW(FALSE);
    DestroyTimers();
    ReleaseRebarSubclass();
    if(m_hwndHost != nullptr && ::IsWindow(m_hwndHost)) {
        ::DestroyWindow(m_hwndHost);
        m_hwndHost = nullptr;
    }
    if(m_hWnd != nullptr && ::IsWindow(m_hWnd)) {
        ::DestroyWindow(m_hWnd);
        m_hWnd = nullptr;
    }
    if(m_spExplorer) {
        m_spExplorer.Release();
    }
    if(m_spInputObjectSite) {
        m_spInputObjectSite.Release();
    }
    if(m_spServiceProvider) {
        m_spServiceProvider.Release();
    }
    m_spSite.Release();
    m_hwndRebar = nullptr;
    return S_OK;
}

IFACEMETHODIMP QTTabBarClass::ResizeBorderDW(LPCRECT /*prcBorder*/, IUnknown* /*punkToolbarSite*/, BOOL /*fReserved*/) {
    return E_NOTIMPL;
}

IFACEMETHODIMP QTTabBarClass::GetBandInfo(DWORD dwBandID, DWORD dwViewMode, DESKBANDINFO* pdbi) {
    if(pdbi == nullptr) {
        return E_POINTER;
    }

    m_bandId = dwBandID;
    m_vertical = (dwViewMode & DBIF_VIEWMODE_VERTICAL) != 0;

    if((pdbi->dwMask & DBIM_MINSIZE) != 0) {
        pdbi->ptMinSize.x = m_minSize.cx;
        pdbi->ptMinSize.y = m_minSize.cy;
    }
    if((pdbi->dwMask & DBIM_MAXSIZE) != 0) {
        pdbi->ptMaxSize.x = m_maxSize.cx;
        pdbi->ptMaxSize.y = m_maxSize.cy;
    }
    if((pdbi->dwMask & DBIM_ACTUAL) != 0) {
        pdbi->ptActual.x = m_minSize.cx;
        pdbi->ptActual.y = m_minSize.cy;
    }
    if((pdbi->dwMask & DBIM_INTEGRAL) != 0) {
        pdbi->ptIntegral.x = m_minSize.cx;
        pdbi->ptIntegral.y = 0;
    }
    if((pdbi->dwMask & DBIM_MODEFLAGS) != 0) {
        pdbi->dwModeFlags = DBIMF_NORMAL | DBIMF_USECHEVRON;
    }
    if((pdbi->dwMask & DBIM_BKCOLOR) != 0) {
        pdbi->dwMask &= ~DBIM_BKCOLOR;
    }
    if((pdbi->dwMask & DBIM_TITLE) != 0) {
        ::StringCchCopyW(pdbi->wszTitle, ARRAYSIZE(pdbi->wszTitle), L"QTTabBar");
    }

    return S_OK;
}

IFACEMETHODIMP QTTabBarClass::UIActivateIO(BOOL fActivate, MSG* /*pMsg*/) {
    if(fActivate) {
        EnsureWindow();
        if(m_hWnd != nullptr) {
            ::SetFocus(m_hWnd);
        }
    }
    return S_OK;
}

IFACEMETHODIMP QTTabBarClass::HasFocusIO() {
    if(m_hWnd == nullptr) {
        return S_FALSE;
    }
    HWND focus = ::GetFocus();
    if(focus == m_hWnd || ::IsChild(m_hWnd, focus)) {
        return S_OK;
    }
    return S_FALSE;
}

IFACEMETHODIMP QTTabBarClass::TranslateAcceleratorIO(MSG* pMsg) {
    if(pMsg == nullptr) {
        return E_POINTER;
    }
    if(pMsg->message == WM_KEYDOWN && (pMsg->wParam == VK_TAB || pMsg->wParam == VK_F6)) {
        if(m_hWnd != nullptr) {
            ::SetFocus(m_hWnd);
            return S_OK;
        }
    }
    return S_FALSE;
}

IFACEMETHODIMP QTTabBarClass::SetSite(IUnknown* pUnkSite) {
    if(pUnkSite == nullptr) {
        m_spInputObjectSite.Release();
        m_spServiceProvider.Release();
        m_spExplorer.Release();
        m_spSite.Release();
        m_hwndRebar = nullptr;
        ReleaseRebarSubclass();
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

    CComPtr<IOleWindow> spOleWindow;
    if(SUCCEEDED(pUnkSite->QueryInterface(IID_PPV_ARGS(&spOleWindow)))) {
        spOleWindow->GetWindow(&m_hwndRebar);
    }

    EnsureRebarSubclass();
    return S_OK;
}

IFACEMETHODIMP QTTabBarClass::GetSite(REFIID riid, void** ppvSite) {
    if(ppvSite == nullptr) {
        return E_POINTER;
    }
    if(!m_spSite) {
        *ppvSite = nullptr;
        return E_FAIL;
    }
    return m_spSite->QueryInterface(riid, ppvSite);
}

IFACEMETHODIMP QTTabBarClass::GetClassID(CLSID* pClassID) {
    if(pClassID == nullptr) {
        return E_POINTER;
    }
    *pClassID = CLSID_QTTabBarClass;
    return S_OK;
}

IFACEMETHODIMP QTTabBarClass::IsDirty() {
    return S_FALSE;
}

IFACEMETHODIMP QTTabBarClass::Load(IStream* /*pStm*/) {
    return S_OK;
}

IFACEMETHODIMP QTTabBarClass::Save(IStream* /*pStm*/, BOOL /*fClearDirty*/) {
    return S_OK;
}

IFACEMETHODIMP QTTabBarClass::GetSizeMax(ULARGE_INTEGER* pcbSize) {
    if(pcbSize == nullptr) {
        return E_POINTER;
    }
    pcbSize->QuadPart = 0;
    return E_NOTIMPL;
}
