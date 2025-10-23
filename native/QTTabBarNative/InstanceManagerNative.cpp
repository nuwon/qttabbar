#include "pch.h"
#include "InstanceManagerNative.h"

#include "QTButtonBar.h"
#include "QTTabBarClass.h"

InstanceManagerNative& InstanceManagerNative::Instance() {
    static InstanceManagerNative instance;
    return instance;
}

void InstanceManagerNative::RegisterTabBar(HWND explorerHwnd, QTTabBarClass* tabBar) {
    if(explorerHwnd == nullptr || tabBar == nullptr) {
        return;
    }
    std::scoped_lock lock(mutex_);
    Entry& entry = map_[explorerHwnd];
    entry.tabBar = tabBar;
}

void InstanceManagerNative::UnregisterTabBar(QTTabBarClass* tabBar) {
    if(tabBar == nullptr) {
        return;
    }
    std::scoped_lock lock(mutex_);
    for(auto it = map_.begin(); it != map_.end();) {
        if(it->second.tabBar == tabBar) {
            it->second.tabBar = nullptr;
            if(it->second.buttonBar == nullptr) {
                it = map_.erase(it);
                continue;
            }
        }
        ++it;
    }
}

void InstanceManagerNative::RegisterButtonBar(HWND explorerHwnd, QTButtonBar* buttonBar) {
    if(explorerHwnd == nullptr || buttonBar == nullptr) {
        return;
    }
    std::scoped_lock lock(mutex_);
    Entry& entry = map_[explorerHwnd];
    entry.buttonBar = buttonBar;
}

void InstanceManagerNative::UnregisterButtonBar(QTButtonBar* buttonBar) {
    if(buttonBar == nullptr) {
        return;
    }
    std::scoped_lock lock(mutex_);
    for(auto it = map_.begin(); it != map_.end();) {
        if(it->second.buttonBar == buttonBar) {
            it->second.buttonBar = nullptr;
            if(it->second.tabBar == nullptr) {
                it = map_.erase(it);
                continue;
            }
        }
        ++it;
    }
}

QTTabBarClass* InstanceManagerNative::FindTabBar(HWND explorerHwnd) const {
    if(explorerHwnd == nullptr) {
        return nullptr;
    }
    std::scoped_lock lock(mutex_);
    auto it = map_.find(explorerHwnd);
    if(it == map_.end()) {
        return nullptr;
    }
    return it->second.tabBar;
}

QTButtonBar* InstanceManagerNative::FindButtonBar(HWND explorerHwnd) const {
    if(explorerHwnd == nullptr) {
        return nullptr;
    }
    std::scoped_lock lock(mutex_);
    auto it = map_.find(explorerHwnd);
    if(it == map_.end()) {
        return nullptr;
    }
    return it->second.buttonBar;
}

void InstanceManagerNative::NotifyButtonCommand(HWND explorerHwnd, UINT commandId) {
    if(auto* tabBar = FindTabBar(explorerHwnd); tabBar != nullptr) {
        tabBar->HandleButtonCommand(commandId);
    }
}

