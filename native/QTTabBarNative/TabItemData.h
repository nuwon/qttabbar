#pragma once

#include <shlobj_core.h>

#include <deque>
#include <memory>
#include <string>

class TabItemData final {
public:
    struct PidlDeleter {
        void operator()(ITEMIDLIST* pidl) const noexcept {
            if(pidl != nullptr) {
                ::CoTaskMemFree(pidl);
            }
        }
    };

    using UniquePidl = std::unique_ptr<ITEMIDLIST, PidlDeleter>;
    using HistoryStack = std::deque<UniquePidl>;

    TabItemData();
    explicit TabItemData(PCIDLIST_ABSOLUTE pidl);
    TabItemData(const TabItemData& other);
    TabItemData& operator=(const TabItemData& other);
    TabItemData(TabItemData&& other) noexcept = default;
    TabItemData& operator=(TabItemData&& other) noexcept = default;
    ~TabItemData();

    void SetText(std::wstring text);
    const std::wstring& GetText() const noexcept;

    void SetTooltip(std::wstring tooltip);
    const std::wstring& GetTooltip() const noexcept;

    void SetPidl(PCIDLIST_ABSOLUTE pidl);
    PCIDLIST_ABSOLUTE GetPidl() const noexcept;

    void PushBackHistory(PCIDLIST_ABSOLUTE pidl);
    void PushForwardHistory(PCIDLIST_ABSOLUTE pidl);
    UniquePidl PopBackHistory();
    UniquePidl PopForwardHistory();
    void ClearHistory();

    const HistoryStack& GetBackHistory() const noexcept;
    const HistoryStack& GetForwardHistory() const noexcept;

    void SetBounds(const RECT& bounds) noexcept;
    const RECT& GetBounds() const noexcept;

    void SetCloseButtonBounds(const RECT& bounds) noexcept;
    const RECT& GetCloseButtonBounds() const noexcept;

    void SetPreferredSize(SIZE size) noexcept;
    SIZE GetPreferredSize() const noexcept;

    void SetSelected(bool selected) noexcept;
    bool IsSelected() const noexcept;

    void SetHot(bool hot) noexcept;
    bool IsHot() const noexcept;

    void SetClosable(bool closable) noexcept;
    bool IsClosable() const noexcept;

    void SetDirty(bool dirty) noexcept;
    bool IsDirty() const noexcept;

    void SetRowIndex(int index) noexcept;
    int GetRowIndex() const noexcept;

    static UniquePidl ClonePidl(PCIDLIST_ABSOLUTE pidl);

private:
    UniquePidl CloneUnique(const UniquePidl& other) const;

    UniquePidl m_pidl;
    HistoryStack m_backHistory;
    HistoryStack m_forwardHistory;
    std::wstring m_text;
    std::wstring m_tooltip;
    RECT m_bounds;
    RECT m_closeButtonBounds;
    SIZE m_preferredSize;
    bool m_selected;
    bool m_hot;
    bool m_closable;
    bool m_dirty;
    int m_rowIndex;
};

