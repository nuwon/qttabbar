#pragma once

class QTTabBarClass;
class QTButtonBar;

namespace qttabbar {

// Lightweight coordination helper mirroring the managed InstanceManager behavior.
class InstanceManagerNative final {
public:
    static void RegisterTabBar(QTTabBarClass* tabBar) noexcept;
    static void UnregisterTabBar(QTTabBarClass* tabBar) noexcept;
    static QTTabBarClass* GetActiveTabBar() noexcept;

    static void RegisterButtonBar(QTButtonBar* buttonBar) noexcept;
    static void UnregisterButtonBar(QTButtonBar* buttonBar) noexcept;
    static QTButtonBar* GetActiveButtonBar() noexcept;
};

}  // namespace qttabbar

