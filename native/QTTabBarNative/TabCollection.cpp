#include "pch.h"
#include "TabCollection.h"

TabCollection::TabCollection()
    : m_selectedIndex(static_cast<size_t>(-1)) {
}

size_t TabCollection::size() const noexcept {
    return m_tabs.size();
}

bool TabCollection::empty() const noexcept {
    return m_tabs.empty();
}

TabItemData* TabCollection::at(size_t index) noexcept {
    if(index >= m_tabs.size()) {
        return nullptr;
    }
    return m_tabs[index].get();
}

const TabItemData* TabCollection::at(size_t index) const noexcept {
    if(index >= m_tabs.size()) {
        return nullptr;
    }
    return m_tabs[index].get();
}

size_t TabCollection::Add(TabPtr tab) {
    ATLASSERT(tab != nullptr);
    m_tabs.emplace_back(std::move(tab));
    return m_tabs.size() - 1;
}

void TabCollection::Insert(size_t index, TabPtr tab) {
    if(index > m_tabs.size()) {
        index = m_tabs.size();
    }
    m_tabs.insert(m_tabs.begin() + static_cast<ptrdiff_t>(index), std::move(tab));
    if(m_selectedIndex >= index && m_selectedIndex != static_cast<size_t>(-1)) {
        ++m_selectedIndex;
    }
}

void TabCollection::Remove(size_t index) {
    if(index >= m_tabs.size()) {
        return;
    }
    m_tabs.erase(m_tabs.begin() + static_cast<ptrdiff_t>(index));
    if(m_tabs.empty()) {
        m_selectedIndex = static_cast<size_t>(-1);
        return;
    }
    if(m_selectedIndex == index) {
        m_selectedIndex = std::min(index, m_tabs.size() - 1);
    } else if(m_selectedIndex > index && m_selectedIndex != static_cast<size_t>(-1)) {
        --m_selectedIndex;
    }
}

void TabCollection::Clear() {
    m_tabs.clear();
    m_selectedIndex = static_cast<size_t>(-1);
}

void TabCollection::SetSelectedIndex(size_t index) {
    if(index >= m_tabs.size()) {
        m_selectedIndex = static_cast<size_t>(-1);
        return;
    }
    if(m_selectedIndex < m_tabs.size()) {
        if(auto* selected = m_tabs[m_selectedIndex].get()) {
            selected->SetSelected(false);
        }
    }
    m_selectedIndex = index;
    if(auto* selected = m_tabs[m_selectedIndex].get()) {
        selected->SetSelected(true);
    }
}

size_t TabCollection::GetSelectedIndex() const noexcept {
    return m_selectedIndex;
}

TabItemData* TabCollection::GetSelected() noexcept {
    if(m_selectedIndex >= m_tabs.size()) {
        return nullptr;
    }
    return m_tabs[m_selectedIndex].get();
}

const TabItemData* TabCollection::GetSelected() const noexcept {
    if(m_selectedIndex >= m_tabs.size()) {
        return nullptr;
    }
    return m_tabs[m_selectedIndex].get();
}

int TabCollection::IndexOf(const TabItemData* tab) const noexcept {
    for(size_t i = 0; i < m_tabs.size(); ++i) {
        if(m_tabs[i].get() == tab) {
            return static_cast<int>(i);
        }
    }
    return -1;
}

