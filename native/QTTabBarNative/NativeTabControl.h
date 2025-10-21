#pragma once

#include "NativeTabNotifications.h"
#include "TabCollection.h"
#include "NativeTabUpDown.h"

#include <gdiplus.h>

#include <memory>

class NativeTabControl final
    : public CWindowImpl<NativeTabControl, CWindow, CControlWinTraits> {
public:
    DECLARE_WND_CLASS_EX(L"QTNativeTabControl", CS_DBLCLKS, COLOR_WINDOW);

    NativeTabControl() noexcept;
    ~NativeTabControl() override;

    BEGIN_MSG_MAP(NativeTabControl)
        MESSAGE_HANDLER(WM_CREATE, OnCreate)
        MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
        MESSAGE_HANDLER(WM_PAINT, OnPaint)
        MESSAGE_HANDLER(WM_PRINTCLIENT, OnPrintClient)
        MESSAGE_HANDLER(WM_ERASEBKGND, OnEraseBackground)
        MESSAGE_HANDLER(WM_SIZE, OnSize)
        MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
        MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
        MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
        MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
        MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
        MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
        MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
        MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
        MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
        MESSAGE_HANDLER(WM_GETDLGCODE, OnGetDlgCode)
        MESSAGE_HANDLER(WM_TIMER, OnTimer)
        MESSAGE_HANDLER(WM_COMMAND, OnCommand)
        MESSAGE_HANDLER(WM_NOTIFY, OnNotify)
        MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
        MESSAGE_HANDLER(WM_DPICHANGED, OnDpiChanged)
    END_MSG_MAP()

    TabCollection& Tabs() noexcept { return m_tabs; }
    const TabCollection& Tabs() const noexcept { return m_tabs; }

    HWND GetTooltipWindow() const noexcept { return m_hwndTooltip; }

    void EnsureLayout();
    void SetKeyboardFocusVisual(bool enabled) noexcept;
    void SetMultiRow(bool enabled) noexcept;
    void SetThemeAware(bool enabled) noexcept;

    void SelectTab(size_t index);
    void SetHoverTracking(bool enabled);
    void SetHotTab(int index);

private:
    struct Metrics {
        int tabHeight;
        int tabMinWidth;
        int tabMaxWidth;
        int tabPaddingX;
        int tabPaddingY;
        int rowSpacing;
        int tabSpacing;
        SIZE closeButtonSize;
    };

    class ScopedGraphics;

    static void EnsureGdiplus();

    void UpdateMetrics(UINT dpi);
    bool LayoutTabs(const RECT& rcClient);
    void DrawBackground(Gdiplus::Graphics& g, const RECT& rcClient);
    void DrawTabs(Gdiplus::Graphics& g);
    void DrawTab(Gdiplus::Graphics& g, size_t index, TabItemData& tab);
    void DrawCloseButton(Gdiplus::Graphics& g, const RECT& bounds, bool hot, bool pressed, bool selected);

    HFONT CreateScaledFont(UINT dpi);
    void DestroyFont();

    void EnsureTooltip();
    void RebuildTooltip();
    void UpdateUpDown(const RECT& rcClient, bool overflow);

    int HitTestTab(POINT pt) const;
    bool HitTestClose(const TabItemData& tab, POINT pt) const;
    void UpdateHotTab(int index);
    void BeginMouseTrack();
    void EndMouseTrack();

    void BeginDrag(int index, POINT screenPos);
    void EndDrag(int index, POINT screenPos, bool cancel);

    void NotifyState(UINT code, int index);
    void NotifyDrag(UINT code, int index, POINT screenPos);

    void HandleTabSelection(int index, bool fromKeyboard);
    void ActivateChevronMenu(POINT screenPos);

    void HandleKeyboardNavigation(UINT vk, bool ctrlDown);

    LRESULT OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnPrintClient(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnEraseBackground(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnMouseMove(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnMouseLeave(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnLButtonDown(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnLButtonUp(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnLButtonDblClk(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnRButtonUp(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnContextMenu(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnKeyDown(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnGetDlgCode(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnTimer(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnCommand(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnNotify(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnSetCursor(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnDpiChanged(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

    TabCollection m_tabs;
    Metrics m_metrics;
    HFONT m_hFont;
    HTHEME m_hTheme;
    HWND m_hwndTooltip;
    UINT m_currentDpi;
    int m_hotIndex;
    int m_pressedIndex;
    int m_closeHotIndex;
    bool m_trackingMouse;
    bool m_keyboardFocusVisual;
    bool m_multiRow;
    bool m_themeAware;
    bool m_dragging;
    int m_dragIndex;
    POINT m_dragAnchor;
    std::unique_ptr<NativeTabUpDown> m_upDown;
    bool m_showUpDown;
};

