#pragma once

#include "QTTabBarNativeGuids.h"
#include "resource.h"

#include <atlbase.h>
#include <atlcom.h>
#include <atlwin.h>
#include <exdisp.h>
#include <commctrl.h>

#include <memory>
#include <vector>

#include "Config.h"

namespace qttabbar {
class InstanceManagerNative;
}  // namespace qttabbar

class ATL_NO_VTABLE QTButtonBar final
    : public CComObjectRootEx<CComMultiThreadModel>
    , public CComCoClass<QTButtonBar, &CLSID_QTButtonBar>
    , public IQTButtonBar
    , public IDeskBand
    , public IInputObject
    , public IObjectWithSite
    , public IPersistStream
    , public IOleWindow
    , public CWindowImpl<QTButtonBar, CWindow, CControlWinTraits> {
public:
    QTButtonBar() noexcept;
    ~QTButtonBar() override;

    DECLARE_REGISTRY_RESOURCEID(IDR_QTBUTTONBAR)
    DECLARE_NOT_AGGREGATABLE(QTButtonBar)
    DECLARE_PROTECT_FINAL_CONSTRUCT()
    DECLARE_WND_CLASS_EX(L"QTButtonBarWindow", CS_DBLCLKS, COLOR_WINDOW)

    BEGIN_COM_MAP(QTButtonBar)
        COM_INTERFACE_ENTRY(IQTButtonBar)
        COM_INTERFACE_ENTRY(IDeskBand)
        COM_INTERFACE_ENTRY2(IDockingWindow, IDeskBand)
        COM_INTERFACE_ENTRY(IInputObject)
        COM_INTERFACE_ENTRY(IObjectWithSite)
        COM_INTERFACE_ENTRY(IOleWindow)
        COM_INTERFACE_ENTRY(IPersist)
        COM_INTERFACE_ENTRY(IPersistStream)
    END_COM_MAP()

    BEGIN_MSG_MAP(QTButtonBar)
        MESSAGE_HANDLER(WM_CREATE, OnCreate)
        MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
        MESSAGE_HANDLER(WM_SIZE, OnSize)
        MESSAGE_HANDLER(WM_COMMAND, OnCommand)
        MESSAGE_HANDLER(WM_NOTIFY, OnNotify)
        MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
        MESSAGE_HANDLER(WM_KILLFOCUS, OnKillFocus)
        CHAIN_MSG_MAP(CWindowImpl<QTButtonBar, CWindow, CControlWinTraits>)
    END_MSG_MAP()

    HRESULT FinalConstruct();
    void FinalRelease();

    // IOleWindow
    IFACEMETHODIMP GetWindow(HWND* phwnd) override;
    IFACEMETHODIMP ContextSensitiveHelp(BOOL fEnterMode) override;

    // IDeskBand / IDockingWindow
    IFACEMETHODIMP ShowDW(BOOL fShow) override;
    IFACEMETHODIMP CloseDW(DWORD dwReserved) override;
    IFACEMETHODIMP ResizeBorderDW(LPCRECT prcBorder, IUnknown* punkToolbarSite, BOOL fReserved) override;
    IFACEMETHODIMP GetBandInfo(DWORD dwBandID, DWORD dwViewMode, DESKBANDINFO* pdbi) override;

    // IInputObject
    IFACEMETHODIMP UIActivateIO(BOOL fActivate, MSG* pMsg) override;
    IFACEMETHODIMP HasFocusIO() override;
    IFACEMETHODIMP TranslateAcceleratorIO(MSG* pMsg) override;

    // IObjectWithSite
    IFACEMETHODIMP SetSite(IUnknown* pUnkSite) override;
    IFACEMETHODIMP GetSite(REFIID riid, void** ppvSite) override;

    // IPersist / IPersistStream
    IFACEMETHODIMP GetClassID(CLSID* pClassID) override;
    IFACEMETHODIMP IsDirty() override;
    IFACEMETHODIMP Load(IStream* pStm) override;
    IFACEMETHODIMP Save(IStream* pStm, BOOL fClearDirty) override;
    IFACEMETHODIMP GetSizeMax(ULARGE_INTEGER* pcbSize) override;

private:
    enum ButtonIndex {
        BII_NAVIGATION_DROPDOWN = -1,
        BII_SEPARATOR = 0,
        BII_NAVIGATION_BACK = 1,
        BII_NAVIGATION_FORWARD = 2,
        BII_GROUP = 3,
        BII_RECENTTAB = 4,
        BII_APPLICATIONLAUNCHER = 5,
        BII_NEWWINDOW = 6,
        BII_CLONE = 7,
        BII_LOCK = 8,
        BII_MISCTOOL = 9,
        BII_TOPMOST = 10,
        BII_CLOSE_CURRENT = 11,
        BII_CLOSE_ALLBUTCURRENT = 12,
        BII_CLOSE_WINDOW = 13,
        BII_CLOSE_LEFT = 14,
        BII_CLOSE_RIGHT = 15,
        BII_GOUPONELEVEL = 16,
        BII_REFRESH_SHELLBROWSER = 17,
        BII_SHELLSEARCH = 18,
        BII_WINDOWOPACITY = 19,
        BII_FILTERBAR = 20,
        BII_OPTION = 21,
        BII_INTERNAL_BUTTON_COUNT = 22,
    };

    enum CommandId : UINT {
        CMD_FIRST = 0x3400,
        CMD_NAV_BACK = CMD_FIRST + 1,
        CMD_NAV_FORWARD,
        CMD_GROUPS,
        CMD_RECENT_TABS,
        CMD_APP_LAUNCHER,
        CMD_NEW_WINDOW,
        CMD_CLONE_TAB,
        CMD_LOCK_TAB,
        CMD_MISC_TOOLS,
        CMD_TOGGLE_TOPMOST,
        CMD_CLOSE_CURRENT,
        CMD_CLOSE_OTHERS,
        CMD_CLOSE_WINDOW,
        CMD_CLOSE_LEFT,
        CMD_CLOSE_RIGHT,
        CMD_GO_UP,
        CMD_REFRESH,
        CMD_SEARCH,
        CMD_WINDOW_OPACITY,
        CMD_FILTER_BAR,
        CMD_OPTIONS,
    };

    struct ButtonDefinition {
        ButtonIndex index;
        CommandId command;
        UINT style;
        UINT stringResId;
    };

    LRESULT OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnCommand(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnNotify(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnSetFocus(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnKillFocus(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

    HRESULT EnsureWindow();
    void DestroyWindowResources();
    void InitializeToolbar();
    void PopulateButtons();
    void AddButton(const ButtonDefinition& def);
    void AddSeparator();
    void LayoutControls();
    void HandleCommand(UINT commandId);
    void ShowDropdownMenu(UINT commandId, const RECT& buttonRect);
    void DispatchToTabHost(UINT commandId);

    CComPtr<IInputObjectSite> m_spInputObjectSite;
    CComPtr<IServiceProvider> m_spServiceProvider;
    CComPtr<IWebBrowser2> m_spExplorer;
    CComPtr<IUnknown> m_spSite;

    HWND m_hwndToolbar;
    HWND m_hwndSearch;
    HIMAGELIST m_hImageList;
    SIZE m_minSize;
    SIZE m_maxSize;
    bool m_visible;
    bool m_hasFocus;
    qttabbar::ConfigData m_config;
};

