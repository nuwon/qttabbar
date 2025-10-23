#pragma once

#include <windows.h>

#include <mutex>
#include <unordered_map>

class QTTabBarClass;
class QTButtonBar;

class InstanceManagerNative {
public:
    static InstanceManagerNative& Instance();

    void RegisterTabBar(HWND explorerHwnd, QTTabBarClass* tabBar);
    void UnregisterTabBar(QTTabBarClass* tabBar);
    void RegisterButtonBar(HWND explorerHwnd, QTButtonBar* buttonBar);
    void UnregisterButtonBar(QTButtonBar* buttonBar);

    QTTabBarClass* FindTabBar(HWND explorerHwnd) const;
    QTButtonBar* FindButtonBar(HWND explorerHwnd) const;

    void NotifyButtonCommand(HWND explorerHwnd, UINT commandId);

private:
    InstanceManagerNative() = default;

    struct Entry {
        QTTabBarClass* tabBar = nullptr;
        QTButtonBar* buttonBar = nullptr;
    };

    mutable std::mutex mutex_;
    std::unordered_map<HWND, Entry> map_;
};

