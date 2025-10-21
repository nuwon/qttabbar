#pragma once

class NativeTabControl;

class NativeTabUpDown final
    : public CWindowImpl<NativeTabUpDown, CWindow, CControlWinTraits> {
public:
    DECLARE_WND_CLASS_EX(L"QTNativeTabUpDown", CS_DBLCLKS, COLOR_WINDOW);

    NativeTabUpDown() noexcept;
    ~NativeTabUpDown() override = default;

    BEGIN_MSG_MAP(NativeTabUpDown)
        MESSAGE_HANDLER(WM_CREATE, OnCreate)
        MESSAGE_HANDLER(WM_PAINT, OnPaint)
        MESSAGE_HANDLER(WM_ERASEBKGND, OnEraseBackground)
        MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
        MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
        MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
        MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
        MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
    END_MSG_MAP()

    void SetOwner(NativeTabControl* owner) noexcept { m_owner = owner; }
    void UpdateMetrics(UINT dpi);
    void Refresh() const;
    SIZE GetButtonSize() const noexcept { return m_buttonSize; }

private:
    enum class HotPart {
        None,
        Left,
        Right,
    };

    void BeginTrack();
    void ResetTrack();
    HotPart HitTest(POINT pt) const;
    void DrawGlyph(HDC hdc, const RECT& bounds, bool left, bool hot, bool pressed) const;
    void SendScrollCommand(bool forward);

    LRESULT OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnEraseBackground(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnMouseMove(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnMouseLeave(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnLButtonDown(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnLButtonUp(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnSetCursor(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

    NativeTabControl* m_owner;
    UINT m_currentDpi;
    HotPart m_hotPart;
    HotPart m_pressedPart;
    bool m_tracking;
    SIZE m_buttonSize;
};

