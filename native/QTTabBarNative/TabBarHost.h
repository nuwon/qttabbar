#pragma once

#include <atlbase.h>
#include <atlwin.h>
#include <atlcom.h>
#include <exdisp.h>

#include <deque>
#include <optional>
#include <string>
#include <vector>

class ITabBarHostOwner {
public:
    virtual ~ITabBarHostOwner() = default;
    virtual HWND GetHostWindow() const noexcept = 0;
    virtual HWND GetHostRebarWindow() const noexcept = 0;
    virtual void NotifyTabHostFocusChange(BOOL hasFocus) = 0;
};

class TabBarHost final
    : public CWindowImpl<TabBarHost, CWindow, CControlWinTraits>
    , public IDispEventSimpleImpl<1, TabBarHost, &DIID_DWebBrowserEvents2> {
public:
    DECLARE_WND_CLASS_EX(L"QTTabBarNative_TabBarHost", CS_DBLCLKS, COLOR_WINDOW);

    explicit TabBarHost(ITabBarHostOwner& owner) noexcept;
    ~TabBarHost() override;

    void Initialize();
    void SetExplorer(IWebBrowser2* browser);
    void ClearExplorer();
    void ExecuteCommand(UINT commandId);
    void ShowContextMenu(const POINT& screenPoint);
    void SaveSessionState() const;
    void RestoreSessionState();
    void OnBandVisibilityChanged(bool visible);
    bool HandleAccelerator(MSG* pMsg);
    bool HasFocus() const noexcept { return m_hasFocus; }
    void OnParentDestroyed();

    std::vector<std::wstring> GetOpenTabs() const;
    std::vector<std::wstring> GetClosedTabHistory() const;
    void ActivateTabByIndex(std::size_t index);
    void RestoreClosedTabByIndex(std::size_t index);
    void CloneActiveTab();
    void CloseAllTabsExceptActive();
    void CloseTabsToLeft();
    void CloseTabsToRight();
    void GoUpOneLevel();
    void NavigateBack();
    void NavigateForward();
    bool OpenCapturedWindow(const std::wstring& path);
    std::wstring GetCurrentPath() const { return m_currentPath; }

    BEGIN_MSG_MAP(TabBarHost)
        MESSAGE_HANDLER(WM_CREATE, OnCreate)
        MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
        MESSAGE_HANDLER(WM_TIMER, OnTimer)
        MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
        MESSAGE_HANDLER(WM_COMMAND, OnCommand)
        MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
        MESSAGE_HANDLER(WM_KILLFOCUS, OnKillFocus)
        MESSAGE_HANDLER(WM_DPICHANGED, OnDpiChanged)
        MESSAGE_HANDLER(WM_GETDLGCODE, OnGetDlgCode)
        MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
        MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
    END_MSG_MAP()

    BEGIN_SINK_MAP(TabBarHost)
        SINK_ENTRY_INFO(1, DIID_DWebBrowserEvents2, DISPID_BEFORENAVIGATE2, OnBeforeNavigate2, &kBeforeNavigate2Info)
        SINK_ENTRY_INFO(1, DIID_DWebBrowserEvents2, DISPID_NAVIGATECOMPLETE2, OnNavigateComplete2, &kNavigateComplete2Info)
        SINK_ENTRY_INFO(1, DIID_DWebBrowserEvents2, DISPID_DOCUMENTCOMPLETE, OnDocumentComplete, &kDocumentCompleteInfo)
        SINK_ENTRY_INFO(1, DIID_DWebBrowserEvents2, DISPID_NAVIGATEERROR, OnNavigateError, &kNavigateErrorInfo)
        SINK_ENTRY_INFO(1, DIID_DWebBrowserEvents2, DISPID_ONQUIT, OnQuit, &kOnQuitInfo)
    END_SINK_MAP()

private:
    struct TabDescriptor {
        std::wstring path;
        bool active = false;
    };

    static constexpr wchar_t kRegistryRoot[] = L"Software\\QTTabBar\\";
    static constexpr wchar_t kTabsValueName[] = L"TabsOnLastClosedWindow";
    static constexpr UINT kMaxClosedHistory = 20;
    static constexpr UINT kSelectTabTimerMs = 5000;
    static constexpr UINT kContextMenuTimerMs = 0x4B0; // 1200ms

    static _ATL_FUNC_INFO kBeforeNavigate2Info;
    static _ATL_FUNC_INFO kNavigateComplete2Info;
    static _ATL_FUNC_INFO kDocumentCompleteInfo;
    static _ATL_FUNC_INFO kNavigateErrorInfo;
    static _ATL_FUNC_INFO kOnQuitInfo;

    LRESULT OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnTimer(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnContextMenu(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnCommand(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnSetFocus(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnKillFocus(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnDpiChanged(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnGetDlgCode(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnKeyDown(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

    void __stdcall OnBeforeNavigate2(IDispatch* pDisp, VARIANT* url, VARIANT* flags, VARIANT* targetFrameName,
                                     VARIANT* postData, VARIANT* headers, VARIANT_BOOL* cancel);
    void __stdcall OnNavigateComplete2(IDispatch* pDisp, VARIANT* url);
    void __stdcall OnDocumentComplete(IDispatch* pDisp, VARIANT* url);
    void __stdcall OnNavigateError(IDispatch* pDisp, VARIANT* url, VARIANT* targetFrameName, VARIANT* statusCode,
                                   VARIANT_BOOL* cancel);
    void __stdcall OnQuit();

    void ConnectBrowserEvents();
    void DisconnectBrowserEvents();
    void StartTimers();
    void StopTimers();
    void EnsureContextMenu();
    void DestroyContextMenu();
    std::wstring VariantToString(const VARIANT* value) const;
    void UpdateActivePath(const std::wstring& path);
    void AddTab(const std::wstring& path, bool makeActive, bool allowDuplicate);
    void ActivateTab(std::size_t index);
    void ActivateNextTab();
    void ActivatePreviousTab();
    void CloseActiveTab();
    void RestoreLastClosed();
    void RefreshExplorer();
    void OpenOptions();
    void TrimClosedHistory();
    void LogTabsState(const wchar_t* source) const;

    ITabBarHostOwner& m_owner;
    CComPtr<IWebBrowser2> m_spBrowser;
    DWORD m_browserCookie;
    HMENU m_hContextMenu;
    std::vector<TabDescriptor> m_tabs;
    std::deque<std::wstring> m_closedHistory;
    std::optional<std::wstring> m_pendingNavigation;
    std::wstring m_currentPath;
    UINT m_dpiX;
    UINT m_dpiY;
    bool m_hasFocus;
    bool m_visible;
};

