#include "pch.h"
#include "TabItemData.h"

#include <Shlwapi.h>

#include <utility>

#pragma comment(lib, "shlwapi.lib")

namespace {

inline TabItemData::UniquePidl CloneIfValid(PCIDLIST_ABSOLUTE pidl) {
    if(pidl == nullptr) {
        return {};
    }
    PIDLIST_ABSOLUTE clone = ::ILCloneFull(pidl);
    return TabItemData::UniquePidl(clone);
}

} // namespace

TabItemData::TabItemData()
    : m_bounds{0, 0, 0, 0}
    , m_closeButtonBounds{0, 0, 0, 0}
    , m_preferredSize{0, 0}
    , m_selected(false)
    , m_hot(false)
    , m_closable(true)
    , m_dirty(false)
    , m_rowIndex(0) {
}

TabItemData::TabItemData(PCIDLIST_ABSOLUTE pidl)
    : TabItemData() {
    SetPidl(pidl);
}

TabItemData::TabItemData(const TabItemData& other)
    : m_pidl(CloneUnique(other.m_pidl))
    , m_text(other.m_text)
    , m_tooltip(other.m_tooltip)
    , m_bounds(other.m_bounds)
    , m_closeButtonBounds(other.m_closeButtonBounds)
    , m_preferredSize(other.m_preferredSize)
    , m_selected(other.m_selected)
    , m_hot(other.m_hot)
    , m_closable(other.m_closable)
    , m_dirty(other.m_dirty)
    , m_rowIndex(other.m_rowIndex) {
    for(const auto& entry : other.m_backHistory) {
        m_backHistory.emplace_back(CloneUnique(entry));
    }
    for(const auto& entry : other.m_forwardHistory) {
        m_forwardHistory.emplace_back(CloneUnique(entry));
    }
}

TabItemData& TabItemData::operator=(const TabItemData& other) {
    if(this == &other) {
        return *this;
    }

    m_pidl = CloneUnique(other.m_pidl);
    m_text = other.m_text;
    m_tooltip = other.m_tooltip;
    m_bounds = other.m_bounds;
    m_closeButtonBounds = other.m_closeButtonBounds;
    m_preferredSize = other.m_preferredSize;
    m_selected = other.m_selected;
    m_hot = other.m_hot;
    m_closable = other.m_closable;
    m_dirty = other.m_dirty;
    m_rowIndex = other.m_rowIndex;

    m_backHistory.clear();
    for(const auto& entry : other.m_backHistory) {
        m_backHistory.emplace_back(CloneUnique(entry));
    }

    m_forwardHistory.clear();
    for(const auto& entry : other.m_forwardHistory) {
        m_forwardHistory.emplace_back(CloneUnique(entry));
    }

    return *this;
}

TabItemData::~TabItemData() = default;

void TabItemData::SetText(std::wstring text) {
    m_text = std::move(text);
}

const std::wstring& TabItemData::GetText() const noexcept {
    return m_text;
}

void TabItemData::SetTooltip(std::wstring tooltip) {
    m_tooltip = std::move(tooltip);
}

const std::wstring& TabItemData::GetTooltip() const noexcept {
    return m_tooltip;
}

void TabItemData::SetPidl(PCIDLIST_ABSOLUTE pidl) {
    m_pidl = CloneIfValid(pidl);
}

PCIDLIST_ABSOLUTE TabItemData::GetPidl() const noexcept {
    return m_pidl.get();
}

void TabItemData::PushBackHistory(PCIDLIST_ABSOLUTE pidl) {
    m_backHistory.emplace_back(CloneIfValid(pidl));
}

void TabItemData::PushForwardHistory(PCIDLIST_ABSOLUTE pidl) {
    m_forwardHistory.emplace_back(CloneIfValid(pidl));
}

TabItemData::UniquePidl TabItemData::PopBackHistory() {
    if(m_backHistory.empty()) {
        return {};
    }
    UniquePidl pidl = std::move(m_backHistory.back());
    m_backHistory.pop_back();
    return pidl;
}

TabItemData::UniquePidl TabItemData::PopForwardHistory() {
    if(m_forwardHistory.empty()) {
        return {};
    }
    UniquePidl pidl = std::move(m_forwardHistory.back());
    m_forwardHistory.pop_back();
    return pidl;
}

void TabItemData::ClearHistory() {
    m_backHistory.clear();
    m_forwardHistory.clear();
}

const TabItemData::HistoryStack& TabItemData::GetBackHistory() const noexcept {
    return m_backHistory;
}

const TabItemData::HistoryStack& TabItemData::GetForwardHistory() const noexcept {
    return m_forwardHistory;
}

void TabItemData::SetBounds(const RECT& bounds) noexcept {
    m_bounds = bounds;
}

const RECT& TabItemData::GetBounds() const noexcept {
    return m_bounds;
}

void TabItemData::SetCloseButtonBounds(const RECT& bounds) noexcept {
    m_closeButtonBounds = bounds;
}

const RECT& TabItemData::GetCloseButtonBounds() const noexcept {
    return m_closeButtonBounds;
}

void TabItemData::SetPreferredSize(SIZE size) noexcept {
    m_preferredSize = size;
}

SIZE TabItemData::GetPreferredSize() const noexcept {
    return m_preferredSize;
}

void TabItemData::SetSelected(bool selected) noexcept {
    m_selected = selected;
}

bool TabItemData::IsSelected() const noexcept {
    return m_selected;
}

void TabItemData::SetHot(bool hot) noexcept {
    m_hot = hot;
}

bool TabItemData::IsHot() const noexcept {
    return m_hot;
}

void TabItemData::SetClosable(bool closable) noexcept {
    m_closable = closable;
}

bool TabItemData::IsClosable() const noexcept {
    return m_closable;
}

void TabItemData::SetDirty(bool dirty) noexcept {
    m_dirty = dirty;
}

bool TabItemData::IsDirty() const noexcept {
    return m_dirty;
}

void TabItemData::SetRowIndex(int index) noexcept {
    m_rowIndex = index;
}

int TabItemData::GetRowIndex() const noexcept {
    return m_rowIndex;
}

TabItemData::UniquePidl TabItemData::ClonePidl(PCIDLIST_ABSOLUTE pidl) {
    return CloneIfValid(pidl);
}

TabItemData::UniquePidl TabItemData::CloneUnique(const UniquePidl& other) const {
    if(!other) {
        return {};
    }
    return CloneIfValid(other.get());
}

