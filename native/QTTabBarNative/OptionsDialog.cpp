#include "pch.h"
#include "OptionsDialog.h"

#include <CommCtrl.h>

#include <atomic>
#include <memory>
#include <mutex>
#include <optional>
#include <vector>

#include "Config.h"
#include "resource.h"

#pragma comment(lib, "Comctl32.lib")

namespace qttabbar {
namespace {

constexpr UINT kOptionsDialogCreated = WM_APP + 42;
constexpr UINT kOptionsDialogDestroy = WM_APP + 43;
constexpr UINT kOptionsDialogBringToFront = WM_APP + 44;

class OptionsDialogImpl;

class OptionsDialogPage {
public:
    OptionsDialogPage(UINT dialogId, const wchar_t* title)
        : m_dialogId(dialogId)
        , m_title(title) {}

    virtual ~OptionsDialogPage() = default;

    UINT DialogId() const noexcept { return m_dialogId; }
    const wchar_t* Title() const noexcept { return m_title; }
    HWND Hwnd() const noexcept { return m_hwnd; }

    void Create(HWND parent, OptionsDialogImpl* owner) {
        ATLASSERT(m_hwnd == nullptr);
        m_hwnd = ::CreateDialogParamW(
            _AtlBaseModule.GetModuleInstance(),
            MAKEINTRESOURCEW(m_dialogId),
            parent,
            &OptionsDialogPage::DialogProc,
            reinterpret_cast<LPARAM>(this));
        if(m_hwnd != nullptr) {
            ::SetWindowLongPtrW(m_hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(owner));
            OnInitDialog(owner);
        }
    }

    void Destroy() {
        if(m_hwnd != nullptr) {
            ::DestroyWindow(m_hwnd);
            m_hwnd = nullptr;
        }
    }

    void Show(bool visible) {
        if(m_hwnd != nullptr) {
            ::ShowWindow(m_hwnd, visible ? SW_SHOW : SW_HIDE);
        }
    }

    virtual void SyncFromConfig(const ConfigData& config) = 0;
    virtual void SyncToConfig(ConfigData& config) const = 0;

protected:
    virtual void OnInitDialog(OptionsDialogImpl* /*owner*/) {}
    virtual BOOL OnCommand(WORD /*notifyCode*/, WORD /*commandId*/, HWND /*control*/, OptionsDialogImpl* /*owner*/) {
        return FALSE;
    }

private:
    static INT_PTR CALLBACK DialogProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
        if(msg == WM_INITDIALOG) {
            auto* page = reinterpret_cast<OptionsDialogPage*>(lParam);
            ::SetWindowLongPtrW(hwnd, DWLP_USER, reinterpret_cast<LONG_PTR>(page));
            page->m_hwnd = hwnd;
            return TRUE;
        }

        auto* page = reinterpret_cast<OptionsDialogPage*>(::GetWindowLongPtrW(hwnd, DWLP_USER));
        if(page == nullptr) {
            return FALSE;
        }

        switch(msg) {
        case WM_COMMAND: {
            WORD notifyCode = HIWORD(wParam);
            WORD commandId = LOWORD(wParam);
            HWND control = reinterpret_cast<HWND>(lParam);
            auto* owner = reinterpret_cast<OptionsDialogImpl*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));
            if(owner != nullptr) {
                return page->OnCommand(notifyCode, commandId, control, owner);
            }
            break;
        }
        default:
            break;
        }
        return FALSE;
    }

    UINT m_dialogId;
    const wchar_t* m_title;
    HWND m_hwnd = nullptr;
};

class WindowOptionsPage final : public OptionsDialogPage {
public:
    WindowOptionsPage()
        : OptionsDialogPage(IDD_OPTIONS_WINDOW, L"Window") {}

    void SyncFromConfig(const ConfigData& config) override {
        SetCheckbox(IDC_CAPTURE_NEW_WINDOWS, config.window.captureNewWindows);
        SetCheckbox(IDC_CAPTURE_WECHAT_SELECTION, config.window.captureWeChatSelection);
        SetCheckbox(IDC_RESTORE_SESSION, config.window.restoreSession);
        SetCheckbox(IDC_RESTORE_ONLY_LOCKED, config.window.restoreOnlyLocked);
        SetCheckbox(IDC_CLOSE_BTN_UNLOCKED, config.window.closeBtnClosesUnlocked);
        SetCheckbox(IDC_CLOSE_BTN_SINGLE, config.window.closeBtnClosesSingleTab);
        SetCheckbox(IDC_TRAY_ON_CLOSE, config.window.trayOnClose);
        SetCheckbox(IDC_TRAY_ON_MINIMIZE, config.window.trayOnMinimize);
        SetCheckbox(IDC_AUTO_HOOK_WINDOW, config.window.autoHookWindow);
        SetCheckbox(IDC_SHOW_FAIL_NAV_MSG, config.window.showFailNavMsg);
    }

    void SyncToConfig(ConfigData& config) const override {
        config.window.captureNewWindows = IsChecked(IDC_CAPTURE_NEW_WINDOWS);
        config.window.captureWeChatSelection = IsChecked(IDC_CAPTURE_WECHAT_SELECTION);
        config.window.restoreSession = IsChecked(IDC_RESTORE_SESSION);
        config.window.restoreOnlyLocked = IsChecked(IDC_RESTORE_ONLY_LOCKED);
        config.window.closeBtnClosesUnlocked = IsChecked(IDC_CLOSE_BTN_UNLOCKED);
        config.window.closeBtnClosesSingleTab = IsChecked(IDC_CLOSE_BTN_SINGLE);
        config.window.trayOnClose = IsChecked(IDC_TRAY_ON_CLOSE);
        config.window.trayOnMinimize = IsChecked(IDC_TRAY_ON_MINIMIZE);
        config.window.autoHookWindow = IsChecked(IDC_AUTO_HOOK_WINDOW);
        config.window.showFailNavMsg = IsChecked(IDC_SHOW_FAIL_NAV_MSG);
    }

protected:
    void OnInitDialog(OptionsDialogImpl* /*owner*/) override {
        HFONT font = reinterpret_cast<HFONT>(::SendMessageW(GetParent(m_hwnd), WM_GETFONT, 0, 0));
        if(font != nullptr) {
            ::SendMessageW(m_hwnd, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
        }
    }

private:
    void SetCheckbox(int controlId, bool value) {
        if(Hwnd() != nullptr) {
            ::SendDlgItemMessageW(Hwnd(), controlId, BM_SETCHECK, value ? BST_CHECKED : BST_UNCHECKED, 0);
        }
    }

    bool IsChecked(int controlId) const {
        if(Hwnd() != nullptr) {
            return ::SendDlgItemMessageW(Hwnd(), controlId, BM_GETCHECK, 0, 0) == BST_CHECKED;
        }
        return false;
    }
};

class PlaceholderOptionsPage final : public OptionsDialogPage {
public:
    PlaceholderOptionsPage(UINT dialogId, const wchar_t* title)
        : OptionsDialogPage(dialogId, title) {}

    void SyncFromConfig(const ConfigData& /*config*/) override {}
    void SyncToConfig(ConfigData& /*config*/) const override {}
};

class OptionsDialogImpl {
public:
    explicit OptionsDialogImpl(HWND explorerWindow)
        : m_explorerWindow(explorerWindow) {
        m_originalConfig = LoadConfigFromRegistry();
        m_workingConfig = m_originalConfig;
    }

    ~OptionsDialogImpl() {
        for(auto& page : m_pages) {
            page->Destroy();
        }
    }

    HWND hwnd() const noexcept { return m_hwnd; }

    void CreateDialogWindow() {
        ATLASSERT(m_hwnd == nullptr);
        HWND owner = m_explorerWindow;
        m_hwnd = ::CreateDialogParamW(
            _AtlBaseModule.GetModuleInstance(),
            MAKEINTRESOURCEW(IDD_OPTIONS_DIALOG),
            owner,
            &OptionsDialogImpl::DialogProc,
            reinterpret_cast<LPARAM>(this));
        if(m_hwnd != nullptr) {
            ::SetWindowLongPtrW(m_hwnd, DWLP_USER, reinterpret_cast<LONG_PTR>(this));
        }
    }

    void RunMessageLoop() {
        MSG msg;
        while(::GetMessageW(&msg, nullptr, 0, 0)) {
            if(m_hwnd == nullptr || !::IsDialogMessageW(m_hwnd, &msg)) {
                ::TranslateMessage(&msg);
                ::DispatchMessageW(&msg);
            }
        }
    }

    void Show() {
        if(m_hwnd != nullptr) {
            ::ShowWindow(m_hwnd, SW_SHOW);
            ::SetForegroundWindow(m_hwnd);
        }
    }

    void Close() {
        if(m_hwnd != nullptr) {
            ::DestroyWindow(m_hwnd);
            m_hwnd = nullptr;
        }
    }

    void OnInitDialog(HWND hwnd) {
        m_hwnd = hwnd;
        ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));
        InitPages();
        SyncPagesFromConfig();
        HWND tab = ::GetDlgItem(m_hwnd, IDC_OPTIONS_TAB);
        if(tab != nullptr) {
            TabCtrl_SetCurSel(tab, 0);
        }
        SelectPage(0);
    }

    void OnDestroyDialog() {
        m_hwnd = nullptr;
        ::PostQuitMessage(0);
    }

    BOOL OnCommand(WORD notifyCode, WORD commandId, HWND control) {
        switch(commandId) {
        case IDOK:
            ApplyChanges(true);
            return TRUE;
        case IDCANCEL:
            CancelChanges();
            return TRUE;
        case IDC_APPLY_BUTTON:
            ApplyChanges(false);
            return TRUE;
        case IDC_RESET_BUTTON:
            ResetCurrentPage();
            return TRUE;
        default:
            break;
        }

        if(control != nullptr && ::IsWindow(control)) {
            HWND page = CurrentPageWindow();
            if(page != nullptr) {
                if(::SendMessageW(page, WM_COMMAND, MAKEWPARAM(commandId, notifyCode), reinterpret_cast<LPARAM>(control))) {
                    return TRUE;
                }
            }
        }
        return FALSE;
    }

    LRESULT OnNotify(int controlId, LPNMHDR header) {
        if(controlId == IDC_OPTIONS_TAB && header->code == TCN_SELCHANGE) {
            int index = TabCtrl_GetCurSel(header->hwndFrom);
            SelectPage(index);
            return TRUE;
        }
        return FALSE;
    }

    void BringToFront() {
        if(m_hwnd != nullptr) {
            if(::IsIconic(m_hwnd)) {
                ::ShowWindow(m_hwnd, SW_RESTORE);
            }
            ::SetForegroundWindow(m_hwnd);
        }
    }

private:
    static INT_PTR CALLBACK DialogProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
        if(msg == WM_INITDIALOG) {
            auto* self = reinterpret_cast<OptionsDialogImpl*>(lParam);
            if(self != nullptr) {
                self->OnInitDialog(hwnd);
                return TRUE;
            }
            return FALSE;
        }

        auto* self = reinterpret_cast<OptionsDialogImpl*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if(self == nullptr) {
            return FALSE;
        }

        switch(msg) {
        case WM_COMMAND:
            return self->OnCommand(HIWORD(wParam), LOWORD(wParam), reinterpret_cast<HWND>(lParam));
        case WM_NOTIFY:
            return self->OnNotify(static_cast<int>(wParam), reinterpret_cast<LPNMHDR>(lParam));
        case WM_CLOSE:
            self->CancelChanges();
            return TRUE;
        case WM_DESTROY:
            self->OnDestroyDialog();
            return TRUE;
        case WM_DPICHANGED:
            self->OnDpiChanged(reinterpret_cast<RECT*>(lParam));
            return TRUE;
        case kOptionsDialogCreated:
            self->Show();
            return TRUE;
        case kOptionsDialogDestroy:
            self->CancelChanges();
            return TRUE;
        case kOptionsDialogBringToFront:
            self->BringToFront();
            return TRUE;
        default:
            break;
        }
        return FALSE;
    }

    void InitPages() {
        HWND tab = ::GetDlgItem(m_hwnd, IDC_OPTIONS_TAB);
        ATLASSERT(tab != nullptr);
        ::SendMessageW(tab, WM_SETFONT, ::SendMessageW(m_hwnd, WM_GETFONT, 0, 0), TRUE);

        m_pages.emplace_back(std::make_unique<WindowOptionsPage>());
        m_pages.emplace_back(std::make_unique<PlaceholderOptionsPage>(IDD_OPTIONS_TABS, L"Tabs"));
        m_pages.emplace_back(std::make_unique<PlaceholderOptionsPage>(IDD_OPTIONS_TWEAKS, L"Tweaks"));
        m_pages.emplace_back(std::make_unique<PlaceholderOptionsPage>(IDD_OPTIONS_TOOLTIPS, L"Tooltips"));
        m_pages.emplace_back(std::make_unique<PlaceholderOptionsPage>(IDD_OPTIONS_GENERAL, L"General"));
        m_pages.emplace_back(std::make_unique<PlaceholderOptionsPage>(IDD_OPTIONS_APPEARANCE, L"Appearance"));
        m_pages.emplace_back(std::make_unique<PlaceholderOptionsPage>(IDD_OPTIONS_MOUSE, L"Mouse"));
        m_pages.emplace_back(std::make_unique<PlaceholderOptionsPage>(IDD_OPTIONS_KEYS, L"Keys"));
        m_pages.emplace_back(std::make_unique<PlaceholderOptionsPage>(IDD_OPTIONS_GROUPS, L"Groups"));
        m_pages.emplace_back(std::make_unique<PlaceholderOptionsPage>(IDD_OPTIONS_APPS, L"Apps"));
        m_pages.emplace_back(std::make_unique<PlaceholderOptionsPage>(IDD_OPTIONS_BUTTONBAR, L"Button Bar"));
        m_pages.emplace_back(std::make_unique<PlaceholderOptionsPage>(IDD_OPTIONS_PLUGINS, L"Plugins"));
        m_pages.emplace_back(std::make_unique<PlaceholderOptionsPage>(IDD_OPTIONS_LANGUAGE, L"Language"));
        m_pages.emplace_back(std::make_unique<PlaceholderOptionsPage>(IDD_OPTIONS_ABOUT, L"About"));

        RECT pageRect = GetPageRect();
        for(size_t index = 0; index < m_pages.size(); ++index) {
            TCITEMW item{};
            item.mask = TCIF_TEXT;
            item.pszText = const_cast<wchar_t*>(m_pages[index]->Title());
            TabCtrl_InsertItem(tab, static_cast<int>(index), &item);
            m_pages[index]->Create(m_hwnd, this);
            if(m_pages[index]->Hwnd() != nullptr) {
                ::SetWindowPos(m_pages[index]->Hwnd(), nullptr, pageRect.left, pageRect.top,
                               pageRect.right - pageRect.left, pageRect.bottom - pageRect.top,
                               SWP_NOZORDER | SWP_NOACTIVATE);
            }
        }
    }

    void SyncPagesFromConfig() {
        for(auto& page : m_pages) {
            page->SyncFromConfig(m_workingConfig);
        }
    }

    void SyncPagesToConfig() {
        for(auto& page : m_pages) {
            page->SyncToConfig(m_workingConfig);
        }
    }

    void SelectPage(int index) {
        if(index < 0 || index >= static_cast<int>(m_pages.size())) {
            return;
        }
        if(m_activePage == index) {
            return;
        }
        if(m_activePage >= 0) {
            m_pages[m_activePage]->Show(false);
        }
        m_activePage = index;
        m_pages[m_activePage]->Show(true);
    }

    HWND CurrentPageWindow() const {
        if(m_activePage < 0 || m_activePage >= static_cast<int>(m_pages.size())) {
            return nullptr;
        }
        return m_pages[m_activePage]->Hwnd();
    }

    void ApplyChanges(bool closeAfterApply) {
        SyncPagesToConfig();
        WriteConfigToRegistry(m_workingConfig, false);
        UpdateConfigSideEffects(m_workingConfig, true);
        m_originalConfig = m_workingConfig;
        if(closeAfterApply) {
            Close();
        }
    }

    void CancelChanges() {
        m_workingConfig = m_originalConfig;
        Close();
    }

    void ResetCurrentPage() {
        if(m_activePage < 0 || m_activePage >= static_cast<int>(m_pages.size())) {
            return;
        }
        m_pages[m_activePage]->SyncFromConfig(m_originalConfig);
        m_pages[m_activePage]->SyncToConfig(m_workingConfig);
    }

    void OnDpiChanged(RECT* suggestedRect) {
        if(suggestedRect != nullptr) {
            ::SetWindowPos(m_hwnd, nullptr, suggestedRect->left, suggestedRect->top,
                           suggestedRect->right - suggestedRect->left,
                           suggestedRect->bottom - suggestedRect->top,
                           SWP_NOZORDER | SWP_NOACTIVATE);
            RECT pageRect = GetPageRect();
            for(auto& page : m_pages) {
                if(page->Hwnd() != nullptr) {
                    ::SetWindowPos(page->Hwnd(), nullptr, pageRect.left, pageRect.top,
                                   pageRect.right - pageRect.left, pageRect.bottom - pageRect.top,
                                   SWP_NOZORDER | SWP_NOACTIVATE);
                }
            }
        }
    }

    RECT GetPageRect() const {
        RECT tabRect{};
        HWND tab = ::GetDlgItem(m_hwnd, IDC_OPTIONS_TAB);
        if(tab == nullptr) {
            return tabRect;
        }
        ::GetClientRect(tab, &tabRect);
        ::TabCtrl_AdjustRect(tab, FALSE, &tabRect);
        POINT points[2] = {{tabRect.left, tabRect.top}, {tabRect.right, tabRect.bottom}};
        ::MapWindowPoints(tab, m_hwnd, points, 2);
        RECT result{points[0].x, points[0].y, points[1].x, points[1].y};
        return result;
    }

    HWND m_explorerWindow;
    HWND m_hwnd = nullptr;
    std::vector<std::unique_ptr<OptionsDialogPage>> m_pages;
    int m_activePage = -1;
    ConfigData m_originalConfig;
    ConfigData m_workingConfig;
};

struct DialogThreadState {
    std::mutex mutex;
    HANDLE thread = nullptr;
    std::atomic<HWND> hwnd{nullptr};
    std::atomic<bool> running{false};
};

DialogThreadState& GetThreadState() {
    static DialogThreadState state;
    return state;
}

DWORD WINAPI DialogThreadProc(LPVOID param) {
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    bool coInitialized = SUCCEEDED(hr);
    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_TAB_CLASSES | ICC_STANDARD_CLASSES | ICC_BAR_CLASSES;
    ::InitCommonControlsEx(&icc);
    std::shared_ptr<OptionsDialogImpl>* initParam = reinterpret_cast<std::shared_ptr<OptionsDialogImpl>*>(param);
    std::shared_ptr<OptionsDialogImpl> impl = *initParam;
    delete initParam;

    impl->CreateDialogWindow();
    HWND hwnd = impl->hwnd();
    if(hwnd != nullptr) {
        GetThreadState().hwnd.store(hwnd, std::memory_order_release);
        ::PostMessageW(hwnd, kOptionsDialogCreated, 0, 0);
    } else {
        GetThreadState().running.store(false, std::memory_order_release);
        if(coInitialized) {
            CoUninitialize();
        }
        return 0;
    }

    impl->RunMessageLoop();
    GetThreadState().hwnd.store(nullptr, std::memory_order_release);
    if(coInitialized) {
        CoUninitialize();
    }
    GetThreadState().running.store(false, std::memory_order_release);
    return 0;
}

void EnsureThreadStarted(HWND explorerWindow) {
    DialogThreadState& state = GetThreadState();
    std::scoped_lock lock(state.mutex);
    if(state.thread != nullptr) {
        DWORD exitCode = 0;
        if(::GetExitCodeThread(state.thread, &exitCode) && exitCode != STILL_ACTIVE) {
            ::CloseHandle(state.thread);
            state.thread = nullptr;
        }
    }
    if(state.running.load(std::memory_order_acquire)) {
        HWND hwnd = state.hwnd.load(std::memory_order_acquire);
        if(hwnd != nullptr) {
            ::PostMessageW(hwnd, kOptionsDialogBringToFront, 0, 0);
        }
        return;
    }

    auto impl = std::make_shared<OptionsDialogImpl>(explorerWindow);
    auto* implPtr = new std::shared_ptr<OptionsDialogImpl>(impl);
    HANDLE thread = ::CreateThread(nullptr, 0, DialogThreadProc, implPtr, 0, nullptr);
    if(thread != nullptr) {
        state.thread = thread;
        state.running.store(true, std::memory_order_release);
    } else {
        delete implPtr;
    }
}

}  // namespace

void OptionsDialog::Open(HWND explorerWindow) {
    EnsureThreadStarted(explorerWindow);
}

void OptionsDialog::ForceClose() {
    DialogThreadState& state = GetThreadState();
    HWND hwnd = state.hwnd.load(std::memory_order_acquire);
    if(hwnd != nullptr) {
        ::PostMessageW(hwnd, kOptionsDialogDestroy, 0, 0);
    }
}

bool OptionsDialog::IsOpen() {
    return GetThreadState().running.load(std::memory_order_acquire) &&
           GetThreadState().hwnd.load(std::memory_order_acquire) != nullptr;
}

}  // namespace qttabbar

