#include "pch.h"
#include "NativeTabControl.h"

#include "NativeTabNotifications.h"
#include "TabItemData.h"
#include "Resource.h"

#include <CommCtrl.h>
#include <Shlwapi.h>
#include <Uxtheme.h>
#include <Vsstyle.h>
#include <Vssym32.h>
#include <strsafe.h>
#include <windowsx.h>

#include <algorithm>
#include <utility>

#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "comctl32.lib")

namespace {

class GdiplusInitializer {
public:
    GdiplusInitializer() noexcept
        : m_token(0) {
        Gdiplus::GdiplusStartupInput input;
        if(Gdiplus::GdiplusStartup(&m_token, &input, nullptr) != Gdiplus::Ok) {
            m_token = 0;
        }
    }

    ~GdiplusInitializer() {
        if(m_token != 0) {
            Gdiplus::GdiplusShutdown(m_token);
        }
    }

    GdiplusInitializer(const GdiplusInitializer&) = delete;
    GdiplusInitializer& operator=(const GdiplusInitializer&) = delete;

private:
    ULONG_PTR m_token;
};

GdiplusInitializer& GdiModule() {
    static GdiplusInitializer instance;
    return instance;
}

inline UINT GetWindowDpi(HWND hwnd) {
    if(hwnd == nullptr) {
        return USER_DEFAULT_SCREEN_DPI;
    }
    static const auto getDpiForWindow = reinterpret_cast<UINT(WINAPI*)(HWND)>(
        ::GetProcAddress(::GetModuleHandleW(L"user32.dll"), "GetDpiForWindow"));
    if(getDpiForWindow != nullptr) {
        return getDpiForWindow(hwnd);
    }
    HDC hdc = ::GetDC(hwnd);
    UINT dpi = USER_DEFAULT_SCREEN_DPI;
    if(hdc != nullptr) {
        dpi = static_cast<UINT>(::GetDeviceCaps(hdc, LOGPIXELSX));
        ::ReleaseDC(hwnd, hdc);
    }
    return dpi;
}

class ScopedFont final {
public:
    ScopedFont(HDC hdc, HFONT font)
        : m_hdc(hdc)
        , m_oldFont(nullptr) {
        if(m_hdc != nullptr && font != nullptr) {
            m_oldFont = static_cast<HFONT>(::SelectObject(m_hdc, font));
        }
    }

    ~ScopedFont() {
        if(m_hdc != nullptr && m_oldFont != nullptr) {
            ::SelectObject(m_hdc, m_oldFont);
        }
    }

    ScopedFont(const ScopedFont&) = delete;
    ScopedFont& operator=(const ScopedFont&) = delete;

private:
    HDC m_hdc;
    HFONT m_oldFont;
};

} // namespace

class NativeTabControl::ScopedGraphics {
public:
    ScopedGraphics(HWND hwnd, HDC hdcTarget, const RECT& rc)
        : m_hwnd(hwnd)
        , m_target(hdcTarget)
        , m_memDC(nullptr)
        , m_bitmap(nullptr)
        , m_oldBitmap(nullptr)
        , m_rc(rc)
        , m_graphics(nullptr) {
        int width = rc.right - rc.left;
        int height = rc.bottom - rc.top;
        if(width <= 0 || height <= 0) {
            return;
        }
        m_memDC = ::CreateCompatibleDC(m_target);
        if(m_memDC == nullptr) {
            return;
        }
        BITMAPINFO info{};
        info.bmiHeader.biSize = sizeof(info.bmiHeader);
        info.bmiHeader.biWidth = width;
        info.bmiHeader.biHeight = -height;
        info.bmiHeader.biPlanes = 1;
        info.bmiHeader.biBitCount = 32;
        info.bmiHeader.biCompression = BI_RGB;
        void* bits = nullptr;
        m_bitmap = ::CreateDIBSection(m_target, &info, DIB_RGB_COLORS, &bits, nullptr, 0);
        if(m_bitmap == nullptr) {
            ::DeleteDC(m_memDC);
            m_memDC = nullptr;
            return;
        }
        m_oldBitmap = static_cast<HBITMAP>(::SelectObject(m_memDC, m_bitmap));
        ::SetBkMode(m_memDC, TRANSPARENT);
        m_graphics = std::make_unique<Gdiplus::Graphics>(m_memDC);
        m_graphics->SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
        m_graphics->SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);
    }

    ~ScopedGraphics() {
        if(m_memDC != nullptr && m_graphics != nullptr) {
            m_graphics->Flush();
            ::BitBlt(m_target, m_rc.left, m_rc.top, m_rc.right - m_rc.left, m_rc.bottom - m_rc.top, m_memDC, 0, 0, SRCCOPY);
        }
        if(m_oldBitmap != nullptr && m_memDC != nullptr) {
            ::SelectObject(m_memDC, m_oldBitmap);
        }
        if(m_bitmap != nullptr) {
            ::DeleteObject(m_bitmap);
        }
        if(m_memDC != nullptr) {
            ::DeleteDC(m_memDC);
        }
    }

    Gdiplus::Graphics* operator->() const noexcept { return m_graphics.get(); }
    Gdiplus::Graphics& Get() const { return *m_graphics; }
    bool IsValid() const noexcept { return m_graphics != nullptr; }

private:
    HWND m_hwnd;
    HDC m_target;
    HDC m_memDC;
    HBITMAP m_bitmap;
    HBITMAP m_oldBitmap;
    RECT m_rc;
    std::unique_ptr<Gdiplus::Graphics> m_graphics;
};

NativeTabControl::NativeTabControl() noexcept
    : m_metrics{}
    , m_hFont(nullptr)
    , m_hTheme(nullptr)
    , m_hwndTooltip(nullptr)
    , m_currentDpi(USER_DEFAULT_SCREEN_DPI)
    , m_hotIndex(-1)
    , m_pressedIndex(-1)
    , m_closeHotIndex(-1)
    , m_trackingMouse(false)
    , m_keyboardFocusVisual(true)
    , m_multiRow(true)
    , m_themeAware(true)
    , m_dragging(false)
    , m_dragIndex(-1)
    , m_showUpDown(false) {
    m_dragAnchor = POINT{0, 0};
}

NativeTabControl::~NativeTabControl() {
    DestroyWindow();
}

void NativeTabControl::EnsureGdiplus() {
    GdiModule();
}

void NativeTabControl::EnsureLayout() {
    RECT rc{};
    if(!::GetClientRect(m_hWnd, &rc)) {
        return;
    }
    bool overflow = LayoutTabs(rc);
    RebuildTooltip();
    UpdateUpDown(rc, overflow);
    ::InvalidateRect(m_hWnd, nullptr, FALSE);
}

void NativeTabControl::SetKeyboardFocusVisual(bool enabled) noexcept {
    m_keyboardFocusVisual = enabled;
}

void NativeTabControl::SetMultiRow(bool enabled) noexcept {
    if(m_multiRow != enabled) {
        m_multiRow = enabled;
        EnsureLayout();
    }
}

void NativeTabControl::SetThemeAware(bool enabled) noexcept {
    if(m_themeAware == enabled) {
        return;
    }
    m_themeAware = enabled;
    if(!m_themeAware && m_hTheme != nullptr) {
        ::CloseThemeData(m_hTheme);
        m_hTheme = nullptr;
    } else if(m_themeAware && m_hTheme == nullptr) {
        m_hTheme = ::OpenThemeData(m_hWnd, L"Tab");
    }
    ::InvalidateRect(m_hWnd, nullptr, TRUE);
}

void NativeTabControl::SelectTab(size_t index) {
    if(index >= m_tabs.size()) {
        return;
    }
    m_tabs.SetSelectedIndex(index);
    NotifyState(NTCN_TABSELECT, static_cast<int>(index));
    EnsureLayout();
}

void NativeTabControl::SetHoverTracking(bool enabled) {
    if(enabled) {
        BeginMouseTrack();
    } else {
        EndMouseTrack();
    }
}

void NativeTabControl::SetHotTab(int index) {
    UpdateHotTab(index);
}

void NativeTabControl::UpdateMetrics(UINT dpi) {
    m_metrics.tabHeight = MulDiv(28, dpi, USER_DEFAULT_SCREEN_DPI);
    m_metrics.tabMinWidth = MulDiv(96, dpi, USER_DEFAULT_SCREEN_DPI);
    m_metrics.tabMaxWidth = MulDiv(220, dpi, USER_DEFAULT_SCREEN_DPI);
    m_metrics.tabPaddingX = MulDiv(12, dpi, USER_DEFAULT_SCREEN_DPI);
    m_metrics.tabPaddingY = MulDiv(6, dpi, USER_DEFAULT_SCREEN_DPI);
    m_metrics.rowSpacing = MulDiv(4, dpi, USER_DEFAULT_SCREEN_DPI);
    m_metrics.tabSpacing = MulDiv(4, dpi, USER_DEFAULT_SCREEN_DPI);
    m_metrics.closeButtonSize.cx = MulDiv(14, dpi, USER_DEFAULT_SCREEN_DPI);
    m_metrics.closeButtonSize.cy = MulDiv(14, dpi, USER_DEFAULT_SCREEN_DPI);
}

bool NativeTabControl::LayoutTabs(const RECT& rcClient) {
    if(m_tabs.empty()) {
        return false;
    }

    HDC hdc = ::GetDC(m_hWnd);
    if(hdc == nullptr) {
        return false;
    }
    ScopedFont scopedFont(hdc, m_hFont);

    SIZE upDownSize{0, 0};
    int upDownWidth = 0;
    if(!m_multiRow) {
        if(m_upDown) {
            upDownSize = m_upDown->GetButtonSize();
            upDownWidth = upDownSize.cx * 2;
        } else {
            upDownWidth = MulDiv(36, static_cast<int>(m_currentDpi), USER_DEFAULT_SCREEN_DPI);
        }
    }

    int availableWidth = (rcClient.right - rcClient.left) - upDownWidth;
    if(availableWidth < 0) {
        availableWidth = 0;
    }
    int x = rcClient.left;
    int y = rcClient.top;
    int rowIndex = 0;
    int maxHeight = rcClient.top;
    bool overflow = false;
    int maxRight = rcClient.left + availableWidth;

    for(size_t i = 0; i < m_tabs.size(); ++i) {
        TabItemData* tab = m_tabs.at(i);
        if(tab == nullptr) {
            continue;
        }
        std::wstring text = tab->GetText();
        if(text.empty()) {
            text = L"(Empty)";
        }
        RECT textRc{0, 0, 0, 0};
        ::DrawTextW(hdc, text.c_str(), static_cast<int>(text.length()), &textRc, DT_SINGLELINE | DT_CALCRECT);
        int textWidth = textRc.right - textRc.left;
        int tabWidth = textWidth + (m_metrics.tabPaddingX * 2);
        tabWidth += tab->IsClosable() ? (m_metrics.closeButtonSize.cx + m_metrics.tabPaddingX) : 0;
        tabWidth = std::clamp(tabWidth, m_metrics.tabMinWidth, m_metrics.tabMaxWidth);
        int tabHeight = m_metrics.tabHeight;
        int tabRight = x + tabWidth;

        if(m_multiRow && x != rcClient.left && tabRight > rcClient.left + availableWidth) {
            x = rcClient.left;
            y += tabHeight + m_metrics.rowSpacing;
            ++rowIndex;
            tabRight = x + tabWidth;
        }

        if(!m_multiRow) {
            if(x >= maxRight) {
                overflow = true;
                tab->SetBounds(RECT{0, 0, 0, 0});
                tab->SetCloseButtonBounds(RECT{0, 0, 0, 0});
                tab->SetPreferredSize(SIZE{tabWidth, tabHeight});
                continue;
            }
            if(tabRight > maxRight) {
                overflow = true;
                tabRight = maxRight;
            }
        }

        RECT bounds{x, y, tabRight, y + tabHeight};
        tab->SetBounds(bounds);
        tab->SetRowIndex(rowIndex);
        maxHeight = std::max(maxHeight, bounds.bottom);

        if(tab->IsClosable()) {
            RECT closeRect{};
            closeRect.right = bounds.right - m_metrics.tabPaddingX;
            closeRect.left = closeRect.right - m_metrics.closeButtonSize.cx;
            closeRect.top = bounds.top + ((tabHeight - m_metrics.closeButtonSize.cy) / 2);
            closeRect.bottom = closeRect.top + m_metrics.closeButtonSize.cy;
            if(closeRect.left < bounds.left) {
                closeRect = RECT{0, 0, 0, 0};
            }
            tab->SetCloseButtonBounds(closeRect);
        } else {
            tab->SetCloseButtonBounds(RECT{0, 0, 0, 0});
        }

        tab->SetPreferredSize(SIZE{tabWidth, tabHeight});
        x = tabRight + m_metrics.tabSpacing;
    }

    ::ReleaseDC(m_hWnd, hdc);

    if(maxHeight > rcClient.bottom) {
        RECT newRect = rcClient;
        newRect.bottom = maxHeight;
        ::InvalidateRect(m_hWnd, &newRect, FALSE);
    }

    return overflow;
}

void NativeTabControl::DrawBackground(Gdiplus::Graphics& g, const RECT& rcClient) {
    Gdiplus::Rect rect(rcClient.left, rcClient.top, rcClient.right - rcClient.left, rcClient.bottom - rcClient.top);
    if(m_themeAware && m_hTheme != nullptr) {
        HDC hdc = g.GetHDC();
        ::DrawThemeBackground(m_hTheme, hdc, TABP_BODY, 0, &rcClient, nullptr);
        g.ReleaseHDC(hdc);
    } else {
        Gdiplus::SolidBrush brush(Gdiplus::Color(255, 240, 240, 240));
        g.FillRectangle(&brush, rect);
    }
}

void NativeTabControl::DrawTabs(Gdiplus::Graphics& g) {
    for(size_t i = 0; i < m_tabs.size(); ++i) {
        if(TabItemData* tab = m_tabs.at(i)) {
            DrawTab(g, i, *tab);
        }
    }
}

void NativeTabControl::DrawTab(Gdiplus::Graphics& g, size_t index, TabItemData& tab) {
    const RECT& bounds = tab.GetBounds();
    if(bounds.right <= bounds.left || bounds.bottom <= bounds.top) {
        return;
    }

    bool isSelected = tab.IsSelected();
    bool isHot = (static_cast<int>(index) == m_hotIndex) || tab.IsHot();

    Gdiplus::RectF rect(static_cast<Gdiplus::REAL>(bounds.left), static_cast<Gdiplus::REAL>(bounds.top),
                        static_cast<Gdiplus::REAL>(bounds.right - bounds.left), static_cast<Gdiplus::REAL>(bounds.bottom - bounds.top));
    Gdiplus::GraphicsPath path;
    const float radius = static_cast<float>(MulDiv(6, m_currentDpi, USER_DEFAULT_SCREEN_DPI));
    path.AddArc(rect.X, rect.Y, radius, radius, 180.0f, 90.0f);
    path.AddArc(rect.X + rect.Width - radius, rect.Y, radius, radius, 270.0f, 90.0f);
    path.AddArc(rect.X + rect.Width - radius, rect.Y + rect.Height - radius, radius, radius, 0.0f, 90.0f);
    path.AddArc(rect.X, rect.Y + rect.Height - radius, radius, radius, 90.0f, 90.0f);
    path.CloseFigure();

    Gdiplus::Color fillColor = isSelected ? Gdiplus::Color(255, 204, 228, 247)
                                          : isHot ? Gdiplus::Color(255, 226, 237, 250)
                                                  : Gdiplus::Color(255, 245, 245, 245);
    Gdiplus::Color borderColor = isSelected ? Gdiplus::Color(255, 64, 120, 192)
                                            : Gdiplus::Color(255, 182, 200, 224);
    Gdiplus::SolidBrush brush(fillColor);
    g.FillPath(&brush, &path);
    Gdiplus::Pen pen(borderColor, 1.0f);
    g.DrawPath(&pen, &path);

    std::wstring text = tab.GetText();
    if(text.empty()) {
        text = L"(Empty)";
    }
    Gdiplus::RectF textRect(rect);
    textRect.X += static_cast<Gdiplus::REAL>(m_metrics.tabPaddingX);
    textRect.Y += static_cast<Gdiplus::REAL>(m_metrics.tabPaddingY);
    textRect.Width -= static_cast<Gdiplus::REAL>(m_metrics.tabPaddingX * 2);
    textRect.Height -= static_cast<Gdiplus::REAL>(m_metrics.tabPaddingY * 2);
    if(tab.IsClosable()) {
        textRect.Width -= static_cast<Gdiplus::REAL>(m_metrics.closeButtonSize.cx + m_metrics.tabPaddingX);
    }

    LOGFONTW lf{};
    if(::GetObjectW(m_hFont, sizeof(lf), &lf) > 0) {
        Gdiplus::Font font(&lf);
        Gdiplus::Color textColor = isSelected ? Gdiplus::Color(255, 32, 32, 32) : Gdiplus::Color(255, 64, 64, 64);
        if(tab.IsDirty()) {
            textColor = Gdiplus::Color(255, 200, 32, 32);
        }
        Gdiplus::SolidBrush textBrush(textColor);
        Gdiplus::StringFormat format;
        format.SetTrimming(Gdiplus::StringTrimmingEllipsisCharacter);
        format.SetAlignment(Gdiplus::StringAlignmentNear);
        format.SetLineAlignment(Gdiplus::StringAlignmentCenter);
        g.DrawString(text.c_str(), static_cast<INT>(text.size()), &font, textRect, &format, &textBrush);
    }

    if(tab.IsClosable()) {
        const RECT& closeRect = tab.GetCloseButtonBounds();
        bool closeHot = (static_cast<int>(index) == m_closeHotIndex);
        bool closePressed = (static_cast<int>(index) == m_pressedIndex) && closeHot;
        DrawCloseButton(g, closeRect, closeHot, closePressed, isSelected);
    }

    if(isSelected && m_keyboardFocusVisual && (::GetFocus() == m_hWnd)) {
        Gdiplus::RectF focusRect = rect;
        focusRect.Inflate(-2.0f, -2.0f);
        Gdiplus::Pen focusPen(Gdiplus::Color(255, 120, 120, 120), 1.0f);
        focusPen.SetDashStyle(Gdiplus::DashStyleDot);
        g.DrawRectangle(&focusPen, focusRect);
    }
}

void NativeTabControl::DrawCloseButton(Gdiplus::Graphics& g, const RECT& bounds, bool hot, bool pressed, bool selected) {
    if(bounds.right <= bounds.left || bounds.bottom <= bounds.top) {
        return;
    }
    Gdiplus::RectF rect(static_cast<Gdiplus::REAL>(bounds.left), static_cast<Gdiplus::REAL>(bounds.top),
                        static_cast<Gdiplus::REAL>(bounds.right - bounds.left), static_cast<Gdiplus::REAL>(bounds.bottom - bounds.top));
    Gdiplus::Color background = hot ? (pressed ? Gdiplus::Color(255, 219, 68, 55) : Gdiplus::Color(255, 232, 185, 172))
                                    : selected ? Gdiplus::Color(0, 0, 0, 0)
                                               : Gdiplus::Color(0, 0, 0, 0);
    if(background.GetAlpha() > 0) {
        Gdiplus::SolidBrush brush(background);
        g.FillEllipse(&brush, rect);
    }
    Gdiplus::Pen pen(hot ? Gdiplus::Color(255, 80, 80, 80) : Gdiplus::Color(255, 120, 120, 120), 2.0f);
    g.DrawLine(&pen, rect.X + 4.0f, rect.Y + 4.0f, rect.GetRight() - 4.0f, rect.GetBottom() - 4.0f);
    g.DrawLine(&pen, rect.GetRight() - 4.0f, rect.Y + 4.0f, rect.X + 4.0f, rect.GetBottom() - 4.0f);
}

HFONT NativeTabControl::CreateScaledFont(UINT dpi) {
    LOGFONTW lf{};
    if(::SystemParametersInfoW(SPI_GETICONTITLELOGFONT, sizeof(lf), &lf, 0) == FALSE) {
        lf.lfHeight = -MulDiv(12, dpi, USER_DEFAULT_SCREEN_DPI);
        lf.lfWeight = FW_NORMAL;
        ::StringCchCopyW(lf.lfFaceName, ARRAYSIZE(lf.lfFaceName), L"Segoe UI");
    } else {
        lf.lfHeight = MulDiv(lf.lfHeight, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
    }
    return ::CreateFontIndirectW(&lf);
}

void NativeTabControl::DestroyFont() {
    if(m_hFont != nullptr) {
        ::DeleteObject(m_hFont);
        m_hFont = nullptr;
    }
}

void NativeTabControl::EnsureTooltip() {
    if(m_hwndTooltip != nullptr && ::IsWindow(m_hwndTooltip)) {
        return;
    }
    m_hwndTooltip = ::CreateWindowExW(WS_EX_TOPMOST, TOOLTIPS_CLASSW, nullptr,
                                      WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP, 0, 0, 0, 0,
                                      m_hWnd, nullptr, _AtlBaseModule.GetModuleInstance(), nullptr);
    if(m_hwndTooltip != nullptr) {
        ::SetWindowPos(m_hwndTooltip, HWND_TOPMOST, 0, 0, 0, 0,
                       SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        ::SendMessageW(m_hwndTooltip, TTM_SETMAXTIPWIDTH, 0, 400);
    }
}

void NativeTabControl::RebuildTooltip() {
    if(m_hwndTooltip == nullptr || !::IsWindow(m_hwndTooltip)) {
        return;
    }
    // Recreate the tooltip window to avoid stale rectangles.
    ::DestroyWindow(m_hwndTooltip);
    m_hwndTooltip = nullptr;
    EnsureTooltip();
    if(m_hwndTooltip == nullptr) {
        return;
    }
    for(size_t i = 0; i < m_tabs.size(); ++i) {
        if(TabItemData* tab = m_tabs.at(i)) {
            TOOLINFOW ti{};
            ti.cbSize = sizeof(ti);
            ti.uFlags = TTF_SUBCLASS;
            ti.hwnd = m_hWnd;
            ti.uId = static_cast<UINT_PTR>(i + 1);
            ti.lpszText = LPSTR_TEXTCALLBACKW;
            ti.rect = tab->GetBounds();
            ::SendMessageW(m_hwndTooltip, TTM_ADDTOOL, 0, reinterpret_cast<LPARAM>(&ti));
        }
    }
}

void NativeTabControl::UpdateUpDown(const RECT& rcClient, bool overflow) {
    if(!m_upDown || m_upDown->m_hWnd == nullptr) {
        return;
    }
    if(m_multiRow || m_tabs.empty()) {
        ::ShowWindow(m_upDown->m_hWnd, SW_HIDE);
        m_showUpDown = false;
        return;
    }

    if(!overflow) {
        ::ShowWindow(m_upDown->m_hWnd, SW_HIDE);
        m_showUpDown = false;
        return;
    }

    SIZE buttonSize = m_upDown->GetButtonSize();
    int buttonWidth = buttonSize.cx * 2;
    int buttonHeight = std::max(m_metrics.tabHeight, buttonSize.cy);
    RECT bounds = rcClient;
    bounds.left = rcClient.right - buttonWidth;
    bounds.bottom = bounds.top + buttonHeight;

    ::SetWindowPos(m_upDown->m_hWnd, nullptr, bounds.left, bounds.top, buttonWidth, buttonHeight,
                   SWP_NOZORDER | SWP_NOACTIVATE);
    ::ShowWindow(m_upDown->m_hWnd, SW_SHOW);
    m_showUpDown = true;
}

int NativeTabControl::HitTestTab(POINT pt) const {
    for(size_t i = 0; i < m_tabs.size(); ++i) {
        if(const TabItemData* tab = m_tabs.at(i)) {
            if(::PtInRect(&tab->GetBounds(), pt)) {
                return static_cast<int>(i);
            }
        }
    }
    return -1;
}

bool NativeTabControl::HitTestClose(const TabItemData& tab, POINT pt) const {
    RECT rc = tab.GetCloseButtonBounds();
    return ::PtInRect(&rc, pt) != FALSE;
}

void NativeTabControl::UpdateHotTab(int index) {
    if(m_hotIndex == index) {
        return;
    }
    m_hotIndex = index;
    for(size_t i = 0; i < m_tabs.size(); ++i) {
        if(TabItemData* tab = m_tabs.at(i)) {
            tab->SetHot(static_cast<int>(i) == m_hotIndex);
        }
    }
    ::InvalidateRect(m_hWnd, nullptr, FALSE);
}

void NativeTabControl::BeginMouseTrack() {
    if(m_trackingMouse) {
        return;
    }
    TRACKMOUSEEVENT track{};
    track.cbSize = sizeof(track);
    track.dwFlags = TME_LEAVE;
    track.hwndTrack = m_hWnd;
    if(::TrackMouseEvent(&track)) {
        m_trackingMouse = true;
    }
}

void NativeTabControl::EndMouseTrack() {
    m_trackingMouse = false;
}

void NativeTabControl::BeginDrag(int index, POINT screenPos) {
    if(index < 0) {
        return;
    }
    m_dragging = true;
    m_dragIndex = index;
    m_dragAnchor = screenPos;
    NotifyDrag(NTCN_BEGIN_DRAG, index, screenPos);
}

void NativeTabControl::EndDrag(int index, POINT screenPos, bool cancel) {
    if(!m_dragging) {
        return;
    }
    m_dragging = false;
    m_dragIndex = -1;
    if(!cancel) {
        NotifyDrag(NTCN_END_DRAG, index, screenPos);
    }
}

void NativeTabControl::NotifyState(UINT code, int index) {
    if(GetParent() == nullptr) {
        return;
    }
    NMTABSTATECHANGE notify{};
    notify.hdr.hwndFrom = m_hWnd;
    notify.hdr.idFrom = ::GetDlgCtrlID(m_hWnd);
    notify.hdr.code = code;
    notify.index = index;
    TabItemData* tab = nullptr;
    if(index >= 0 && static_cast<size_t>(index) < m_tabs.size()) {
        tab = m_tabs.at(static_cast<size_t>(index));
    }
    notify.tab = tab;
    ::SendMessageW(GetParent(), WM_NOTIFY, notify.hdr.idFrom, reinterpret_cast<LPARAM>(&notify));
}

void NativeTabControl::NotifyDrag(UINT code, int index, POINT screenPos) {
    if(GetParent() == nullptr) {
        return;
    }
    NMTABDRAG notify{};
    notify.hdr.hwndFrom = m_hWnd;
    notify.hdr.idFrom = ::GetDlgCtrlID(m_hWnd);
    notify.hdr.code = code;
    notify.index = index;
    notify.screenPos = screenPos;
    ::SendMessageW(GetParent(), WM_NOTIFY, notify.hdr.idFrom, reinterpret_cast<LPARAM>(&notify));
}

void NativeTabControl::HandleTabSelection(int index, bool fromKeyboard) {
    if(index < 0 || static_cast<size_t>(index) >= m_tabs.size()) {
        return;
    }
    m_tabs.SetSelectedIndex(static_cast<size_t>(index));
    NotifyState(NTCN_TABSELECT, index);
    if(fromKeyboard) {
        ::SetFocus(m_hWnd);
    }
    ::InvalidateRect(m_hWnd, nullptr, FALSE);
}

void NativeTabControl::ActivateChevronMenu(POINT screenPos) {
    if(GetParent() == nullptr) {
        return;
    }
    ::SendMessageW(GetParent(), WM_COMMAND, MAKEWPARAM(ID_CONTEXT_OPTIONS, 0), reinterpret_cast<LPARAM>(m_hWnd));
}

void NativeTabControl::HandleKeyboardNavigation(UINT vk, bool ctrlDown) {
    if(m_tabs.empty()) {
        return;
    }
    size_t selectedIndex = m_tabs.GetSelectedIndex();
    if(selectedIndex >= m_tabs.size()) {
        selectedIndex = 0;
    }

    if(vk == VK_LEFT) {
        if(selectedIndex == 0) {
            selectedIndex = m_tabs.size() - 1;
        } else {
            --selectedIndex;
        }
        HandleTabSelection(static_cast<int>(selectedIndex), true);
    } else if(vk == VK_RIGHT) {
        selectedIndex = (selectedIndex + 1) % m_tabs.size();
        HandleTabSelection(static_cast<int>(selectedIndex), true);
    } else if(vk == VK_HOME) {
        HandleTabSelection(0, true);
    } else if(vk == VK_END) {
        HandleTabSelection(static_cast<int>(m_tabs.size() - 1), true);
    } else if(ctrlDown && vk == VK_TAB) {
        bool shiftDown = (::GetKeyState(VK_SHIFT) & 0x8000) != 0;
        if(shiftDown) {
            if(selectedIndex == 0) {
                selectedIndex = m_tabs.size() - 1;
            } else {
                --selectedIndex;
            }
        } else {
            selectedIndex = (selectedIndex + 1) % m_tabs.size();
        }
        HandleTabSelection(static_cast<int>(selectedIndex), true);
    }
}

LRESULT NativeTabControl::OnCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    EnsureGdiplus();
    m_currentDpi = GetWindowDpi(m_hWnd);
    UpdateMetrics(m_currentDpi);
    DestroyFont();
    m_hFont = CreateScaledFont(m_currentDpi);
    ::SetWindowTheme(m_hWnd, L"Explorer", nullptr);
    if(m_themeAware) {
        m_hTheme = ::OpenThemeData(m_hWnd, L"Tab");
    }
    EnsureTooltip();
    m_upDown = std::make_unique<NativeTabUpDown>();
    if(m_upDown) {
        RECT rc{};
        HWND hwndUp = m_upDown->Create(m_hWnd, rc, nullptr, WS_CHILD | WS_CLIPSIBLINGS, 0, ID_TAB_UPDOWN);
        if(hwndUp != nullptr) {
            m_upDown->SetOwner(this);
            m_upDown->UpdateMetrics(m_currentDpi);
            ::ShowWindow(hwndUp, SW_HIDE);
        } else {
            m_upDown.reset();
        }
    }
    ::SetTimer(m_hWnd, 1, 100, nullptr);
    return 0;
}

LRESULT NativeTabControl::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    ::KillTimer(m_hWnd, 1);
    EndMouseTrack();
    if(m_hwndTooltip != nullptr && ::IsWindow(m_hwndTooltip)) {
        ::DestroyWindow(m_hwndTooltip);
    }
    m_hwndTooltip = nullptr;
    if(m_hTheme != nullptr) {
        ::CloseThemeData(m_hTheme);
        m_hTheme = nullptr;
    }
    DestroyFont();
    if(m_upDown) {
        if(m_upDown->m_hWnd != nullptr && ::IsWindow(m_upDown->m_hWnd)) {
            ::DestroyWindow(m_upDown->m_hWnd);
        }
        m_upDown.reset();
    }
    return 0;
}

LRESULT NativeTabControl::OnPaint(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    PAINTSTRUCT ps;
    if(HDC hdc = ::BeginPaint(m_hWnd, &ps)) {
        RECT rc{};
        ::GetClientRect(m_hWnd, &rc);
        ScopedGraphics graphics(m_hWnd, hdc, rc);
        if(graphics.IsValid()) {
            DrawBackground(graphics.Get(), rc);
            DrawTabs(graphics.Get());
        }
        ::EndPaint(m_hWnd, &ps);
    }
    return 0;
}

LRESULT NativeTabControl::OnPrintClient(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    HDC hdc = reinterpret_cast<HDC>(wParam);
    RECT rc{};
    ::GetClientRect(m_hWnd, &rc);
    ScopedGraphics graphics(m_hWnd, hdc, rc);
    if(graphics.IsValid()) {
        DrawBackground(graphics.Get(), rc);
        DrawTabs(graphics.Get());
    }
    return 0;
}

LRESULT NativeTabControl::OnEraseBackground(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    return 1;
}

LRESULT NativeTabControl::OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    RECT rc{0, 0, LOWORD(lParam), HIWORD(lParam)};
    bool overflow = LayoutTabs(rc);
    RebuildTooltip();
    UpdateUpDown(rc, overflow);
    return 0;
}

LRESULT NativeTabControl::OnMouseMove(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    int index = HitTestTab(pt);
    UpdateHotTab(index);
    BeginMouseTrack();
    m_closeHotIndex = -1;
    if(index >= 0) {
        if(TabItemData* tab = m_tabs.at(static_cast<size_t>(index))) {
            if(tab->IsClosable() && HitTestClose(*tab, pt)) {
                m_closeHotIndex = index;
            }
        }
    }
    ::InvalidateRect(m_hWnd, nullptr, FALSE);
    if((wParam & MK_LBUTTON) != 0 && m_pressedIndex >= 0) {
        POINT screenPos = pt;
        ::ClientToScreen(m_hWnd, &screenPos);
        if(!m_dragging) {
            BeginDrag(m_pressedIndex, screenPos);
        } else {
            NotifyDrag(NTCN_DRAGGING, m_dragIndex, screenPos);
        }
    }
    return 0;
}

LRESULT NativeTabControl::OnMouseLeave(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    UpdateHotTab(-1);
    m_closeHotIndex = -1;
    EndMouseTrack();
    ::InvalidateRect(m_hWnd, nullptr, FALSE);
    return 0;
}

LRESULT NativeTabControl::OnLButtonDown(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    int index = HitTestTab(pt);
    if(index >= 0) {
        m_pressedIndex = index;
        if(TabItemData* tab = m_tabs.at(static_cast<size_t>(index))) {
            if(tab->IsClosable() && HitTestClose(*tab, pt)) {
                m_closeHotIndex = index;
            }
        }
        ::SetCapture(m_hWnd);
    }
    ::InvalidateRect(m_hWnd, nullptr, FALSE);
    return 0;
}

LRESULT NativeTabControl::OnLButtonUp(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    int index = HitTestTab(pt);
    bool handled = false;
    POINT screenPt = pt;
    ::ClientToScreen(m_hWnd, &screenPt);
    if(m_dragging) {
        EndDrag(m_dragIndex, screenPt, index < 0);
    }
    if(index >= 0 && index == m_pressedIndex) {
        if(TabItemData* tab = m_tabs.at(static_cast<size_t>(index))) {
            if(tab->IsClosable() && HitTestClose(*tab, pt)) {
                NotifyState(NTCN_TABCLOSE, index);
                ::SendMessageW(GetParent(), WM_COMMAND, MAKEWPARAM(ID_CONTEXT_CLOSETAB, index), reinterpret_cast<LPARAM>(m_hWnd));
                handled = true;
            }
        }
        if(!handled) {
            HandleTabSelection(index, false);
        }
    }
    if(::GetCapture() == m_hWnd) {
        ::ReleaseCapture();
    }
    m_pressedIndex = -1;
    m_dragging = false;
    m_dragIndex = -1;
    ::InvalidateRect(m_hWnd, nullptr, FALSE);
    return 0;
}

LRESULT NativeTabControl::OnLButtonDblClk(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    int index = HitTestTab(pt);
    if(index >= 0) {
        HandleTabSelection(index, false);
    } else {
        ::SendMessageW(GetParent(), WM_COMMAND, MAKEWPARAM(ID_CONTEXT_NEWTAB, 0), reinterpret_cast<LPARAM>(m_hWnd));
    }
    return 0;
}

LRESULT NativeTabControl::OnRButtonUp(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    int index = HitTestTab(pt);
    NotifyState(NTCN_TABCONTEXT, index);
    return 0;
}

LRESULT NativeTabControl::OnContextMenu(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    POINT pt{};
    if(wParam == reinterpret_cast<WPARAM>(m_hWnd)) {
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
        pt.x = GET_X_LPARAM(lParam);
        pt.y = GET_Y_LPARAM(lParam);
    }
    POINT client = pt;
    ::ScreenToClient(m_hWnd, &client);
    int index = HitTestTab(client);
    NotifyState(NTCN_TABCONTEXT, index);
    ::SendMessageW(GetParent(), WM_COMMAND, MAKEWPARAM(ID_CONTEXT_OPTIONS, index), reinterpret_cast<LPARAM>(m_hWnd));
    return 0;
}

LRESULT NativeTabControl::OnKeyDown(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    bool ctrlDown = (::GetKeyState(VK_CONTROL) & 0x8000) != 0;
    HandleKeyboardNavigation(static_cast<UINT>(wParam), ctrlDown);
    return 0;
}

LRESULT NativeTabControl::OnGetDlgCode(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    return DLGC_WANTARROWS | DLGC_WANTTAB | DLGC_WANTCHARS;
}

LRESULT NativeTabControl::OnTimer(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    if(wParam == 1 && ::GetFocus() == m_hWnd) {
        ::InvalidateRect(m_hWnd, nullptr, FALSE);
    }
    return 0;
}

LRESULT NativeTabControl::OnCommand(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = FALSE;
    HWND source = reinterpret_cast<HWND>(lParam);
    UINT command = LOWORD(wParam);
    if(m_upDown && source == m_upDown->m_hWnd) {
        if(command == ID_TAB_SCROLL_LEFT) {
            HandleKeyboardNavigation(VK_LEFT, false);
        } else if(command == ID_TAB_SCROLL_RIGHT) {
            HandleKeyboardNavigation(VK_RIGHT, false);
        }
        bHandled = TRUE;
        return 0;
    }
    if(HWND parent = GetParent(); parent != nullptr && parent != m_hWnd) {
        ::SendMessageW(parent, WM_COMMAND, wParam, lParam);
    }
    return 0;
}

LRESULT NativeTabControl::OnNotify(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = FALSE;
    NMHDR* hdr = reinterpret_cast<NMHDR*>(lParam);
    if(hdr != nullptr && hdr->hwndFrom == m_hwndTooltip) {
        if(hdr->code == TTN_NEEDTEXTW) {
            auto* info = reinterpret_cast<NMTTDISPINFOW*>(hdr);
            size_t index = info->hdr.idFrom > 0 ? info->hdr.idFrom - 1 : 0;
            if(info->hdr.idFrom > 0 && index < m_tabs.size()) {
                if(TabItemData* tab = m_tabs.at(index)) {
                    info->lpszText = const_cast<wchar_t*>(tab->GetTooltip().c_str());
                }
            }
            if(HWND parent = GetParent()) {
                ::SendMessageW(parent, WM_NOTIFY, wParam, lParam);
            }
            bHandled = TRUE;
            return 0;
        }
    }
    return 0;
}

LRESULT NativeTabControl::OnSetCursor(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    if(LOWORD(lParam) == HTCLIENT) {
        POINT pt;
        ::GetCursorPos(&pt);
        ::ScreenToClient(m_hWnd, &pt);
        int index = HitTestTab(pt);
        if(index >= 0) {
            if(TabItemData* tab = m_tabs.at(static_cast<size_t>(index))) {
                if(tab->IsClosable() && HitTestClose(*tab, pt)) {
                    ::SetCursor(::LoadCursorW(nullptr, IDC_HAND));
                    return TRUE;
                }
            }
        }
        ::SetCursor(::LoadCursorW(nullptr, IDC_ARROW));
        return TRUE;
    }
    bHandled = FALSE;
    return FALSE;
}

LRESULT NativeTabControl::OnDpiChanged(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    UINT dpi = HIWORD(wParam);
    m_currentDpi = dpi;
    UpdateMetrics(m_currentDpi);
    DestroyFont();
    m_hFont = CreateScaledFont(m_currentDpi);
    if(m_upDown) {
        m_upDown->UpdateMetrics(m_currentDpi);
    }
    RECT* const rcNew = reinterpret_cast<RECT*>(lParam);
    if(rcNew != nullptr) {
        ::SetWindowPos(m_hWnd, nullptr, rcNew->left, rcNew->top, rcNew->right - rcNew->left, rcNew->bottom - rcNew->top,
                       SWP_NOZORDER | SWP_NOACTIVATE);
    }
    EnsureLayout();
    return 0;
}

