#pragma once

#include "TabItemData.h"

#include <cstddef>
#include <memory>
#include <vector>

class TabCollection final {
public:
    using TabPtr = std::unique_ptr<TabItemData>;
    using Container = std::vector<TabPtr>;

    TabCollection();

    size_t size() const noexcept;
    bool empty() const noexcept;

    TabItemData* at(size_t index) noexcept;
    const TabItemData* at(size_t index) const noexcept;

    size_t Add(TabPtr tab);
    void Insert(size_t index, TabPtr tab);
    void Remove(size_t index);
    void Clear();

    Container::iterator begin() noexcept { return m_tabs.begin(); }
    Container::iterator end() noexcept { return m_tabs.end(); }
    Container::const_iterator begin() const noexcept { return m_tabs.begin(); }
    Container::const_iterator end() const noexcept { return m_tabs.end(); }

    void SetSelectedIndex(size_t index);
    size_t GetSelectedIndex() const noexcept;
    TabItemData* GetSelected() noexcept;
    const TabItemData* GetSelected() const noexcept;

    int IndexOf(const TabItemData* tab) const noexcept;

private:
    Container m_tabs;
    size_t m_selectedIndex;
};

