#include "pch.h"
#include "NativeTabUpDown.h"

#include "NativeTabControl.h"
#include "Resource.h"

#include <CommCtrl.h>
#include <windowsx.h>

NativeTabUpDown::NativeTabUpDown() noexcept
    : m_owner(nullptr)
    , m_currentDpi(USER_DEFAULT_SCREEN_DPI)
    , m_hotPart(HotPart::None)
    , m_pressedPart(HotPart::None)
    , m_tracking(false)
    , m_buttonSize{MulDiv(18, USER_DEFAULT_SCREEN_DPI, USER_DEFAULT_SCREEN_DPI), MulDiv(24, USER_DEFAULT_SCREEN_DPI, USER_DEFAULT_SCREEN_DPI)} {
}

void NativeTabUpDown::UpdateMetrics(UINT dpi) {
    m_currentDpi = dpi == 0 ? USER_DEFAULT_SCREEN_DPI : dpi;
    m_buttonSize.cx = MulDiv(18, static_cast<int>(m_currentDpi), USER_DEFAULT_SCREEN_DPI);
    m_buttonSize.cy = MulDiv(24, static_cast<int>(m_currentDpi), USER_DEFAULT_SCREEN_DPI);
    if(m_hWnd != nullptr) {
        RECT rc{};
        ::GetClientRect(m_hWnd, &rc);
        int height = rc.bottom - rc.top;
        if(height != m_buttonSize.cy) {
            ::SetWindowPos(m_hWnd, nullptr, 0, 0, m_buttonSize.cx * 2, m_buttonSize.cy, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
        Refresh();
    }
}

void NativeTabUpDown::Refresh() const {
    if(m_hWnd != nullptr) {
        ::InvalidateRect(m_hWnd, nullptr, TRUE);
    }
}

void NativeTabUpDown::BeginTrack() {
    if(m_tracking) {
        return;
    }
    TRACKMOUSEEVENT tme{};
    tme.cbSize = sizeof(tme);
    tme.dwFlags = TME_LEAVE;
    tme.hwndTrack = m_hWnd;
    if(::TrackMouseEvent(&tme)) {
        m_tracking = true;
    }
}

void NativeTabUpDown::ResetTrack() {
    m_tracking = false;
    m_hotPart = HotPart::None;
    m_pressedPart = HotPart::None;
}

NativeTabUpDown::HotPart NativeTabUpDown::HitTest(POINT pt) const {
    RECT rc{};
    ::GetClientRect(m_hWnd, &rc);
    LONG mid = (rc.left + rc.right) / 2;
    if(pt.x < rc.left || pt.x > rc.right || pt.y < rc.top || pt.y > rc.bottom) {
        return HotPart::None;
    }
    return pt.x < mid ? HotPart::Left : HotPart::Right;
}

void NativeTabUpDown::DrawGlyph(HDC hdc, const RECT& bounds, bool left, bool hot, bool pressed) const {
    COLORREF base = hot ? RGB(32, 120, 220) : RGB(96, 96, 96);
    if(pressed) {
        base = RGB(20, 90, 180);
    }
    POINT center{(bounds.left + bounds.right) / 2, (bounds.top + bounds.bottom) / 2};
    POINT pts[3];
    if(left) {
        pts[0] = POINT{center.x + MulDiv(-4, m_currentDpi, USER_DEFAULT_SCREEN_DPI), center.y};
        pts[1] = POINT{center.x + MulDiv(2, m_currentDpi, USER_DEFAULT_SCREEN_DPI), center.y - MulDiv(5, m_currentDpi, USER_DEFAULT_SCREEN_DPI)};
        pts[2] = POINT{center.x + MulDiv(2, m_currentDpi, USER_DEFAULT_SCREEN_DPI), center.y + MulDiv(5, m_currentDpi, USER_DEFAULT_SCREEN_DPI)};
    } else {
        pts[0] = POINT{center.x + MulDiv(4, m_currentDpi, USER_DEFAULT_SCREEN_DPI), center.y};
        pts[1] = POINT{center.x + MulDiv(-2, m_currentDpi, USER_DEFAULT_SCREEN_DPI), center.y - MulDiv(5, m_currentDpi, USER_DEFAULT_SCREEN_DPI)};
        pts[2] = POINT{center.x + MulDiv(-2, m_currentDpi, USER_DEFAULT_SCREEN_DPI), center.y + MulDiv(5, m_currentDpi, USER_DEFAULT_SCREEN_DPI)};
    }
    HBRUSH brush = ::CreateSolidBrush(base);
    HPEN pen = ::CreatePen(PS_SOLID, 1, base);
    HGDIOBJ oldBrush = ::SelectObject(hdc, brush);
    HGDIOBJ oldPen = ::SelectObject(hdc, pen);
    ::Polygon(hdc, pts, 3);
    ::SelectObject(hdc, oldBrush);
    ::SelectObject(hdc, oldPen);
    ::DeleteObject(brush);
    ::DeleteObject(pen);
}

void NativeTabUpDown::SendScrollCommand(bool forward) {
    if(m_owner == nullptr || m_owner->m_hWnd == nullptr) {
        return;
    }
    ::SendMessageW(m_owner->m_hWnd, WM_COMMAND,
                   MAKEWPARAM(forward ? ID_TAB_SCROLL_RIGHT : ID_TAB_SCROLL_LEFT, 0),
                   reinterpret_cast<LPARAM>(m_hWnd));
}

LRESULT NativeTabUpDown::OnCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    UpdateMetrics(m_currentDpi);
    return 0;
}

LRESULT NativeTabUpDown::OnPaint(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    PAINTSTRUCT ps;
    if(HDC hdc = ::BeginPaint(m_hWnd, &ps)) {
        RECT rc{};
        ::GetClientRect(m_hWnd, &rc);
        HBRUSH background = ::CreateSolidBrush(RGB(245, 245, 245));
        ::FillRect(hdc, &rc, background);
        ::DeleteObject(background);

        RECT leftRect = rc;
        leftRect.right = (rc.left + rc.right) / 2;
        RECT rightRect = rc;
        rightRect.left = leftRect.right;

        if(m_hotPart == HotPart::Left) {
            HBRUSH hotBrush = ::CreateSolidBrush(RGB(220, 232, 246));
            ::FillRect(hdc, &leftRect, hotBrush);
            ::DeleteObject(hotBrush);
        }
        if(m_hotPart == HotPart::Right) {
            HBRUSH hotBrush = ::CreateSolidBrush(RGB(220, 232, 246));
            ::FillRect(hdc, &rightRect, hotBrush);
            ::DeleteObject(hotBrush);
        }

        DrawGlyph(hdc, leftRect, true, m_hotPart == HotPart::Left, m_pressedPart == HotPart::Left);
        DrawGlyph(hdc, rightRect, false, m_hotPart == HotPart::Right, m_pressedPart == HotPart::Right);

        ::EndPaint(m_hWnd, &ps);
    }
    return 0;
}

LRESULT NativeTabUpDown::OnEraseBackground(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    return 1;
}

LRESULT NativeTabUpDown::OnMouseMove(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    HotPart part = HitTest(pt);
    if(part != m_hotPart) {
        m_hotPart = part;
        Refresh();
    }
    if(part != HotPart::None) {
        BeginTrack();
    }
    return 0;
}

LRESULT NativeTabUpDown::OnMouseLeave(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled) {
    bHandled = TRUE;
    ResetTrack();
    Refresh();
    return 0;
}

LRESULT NativeTabUpDown::OnLButtonDown(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    m_pressedPart = HitTest(pt);
    if(m_pressedPart != HotPart::None) {
        ::SetCapture(m_hWnd);
    }
    Refresh();
    return 0;
}

LRESULT NativeTabUpDown::OnLButtonUp(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    if(::GetCapture() == m_hWnd) {
        ::ReleaseCapture();
    }
    POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    HotPart part = HitTest(pt);
    if(part != HotPart::None && part == m_pressedPart) {
        SendScrollCommand(part == HotPart::Right);
    }
    m_pressedPart = HotPart::None;
    Refresh();
    return 0;
}

LRESULT NativeTabUpDown::OnSetCursor(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
    bHandled = TRUE;
    if(LOWORD(lParam) == HTCLIENT) {
        ::SetCursor(::LoadCursorW(nullptr, IDC_HAND));
        return TRUE;
    }
    bHandled = FALSE;
    return FALSE;
}

