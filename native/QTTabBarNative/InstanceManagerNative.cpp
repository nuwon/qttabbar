#include "pch.h"
#include "InstanceManagerNative.h"

#include "QTTabBarClass.h"
#include "QTButtonBar.h"

namespace qttabbar {
namespace {

thread_local QTTabBarClass* g_activeTabBar = nullptr;
thread_local QTButtonBar* g_activeButtonBar = nullptr;

}  // namespace

void InstanceManagerNative::RegisterTabBar(QTTabBarClass* tabBar) noexcept {
    g_activeTabBar = tabBar;
}

void InstanceManagerNative::UnregisterTabBar(QTTabBarClass* tabBar) noexcept {
    if(g_activeTabBar == tabBar) {
        g_activeTabBar = nullptr;
    }
}

QTTabBarClass* InstanceManagerNative::GetActiveTabBar() noexcept {
    return g_activeTabBar;
}

void InstanceManagerNative::RegisterButtonBar(QTButtonBar* buttonBar) noexcept {
    g_activeButtonBar = buttonBar;
}

void InstanceManagerNative::UnregisterButtonBar(QTButtonBar* buttonBar) noexcept {
    if(g_activeButtonBar == buttonBar) {
        g_activeButtonBar = nullptr;
    }
}

QTButtonBar* InstanceManagerNative::GetActiveButtonBar() noexcept {
    return g_activeButtonBar;
}

}  // namespace qttabbar

