#pragma once

#include <windows.h>

namespace qttabbar::hooks {

constexpr UINT WM_APP_CAPTURE_NEW_WINDOW = WM_APP + 0x120;

struct CaptureNewWindowRequest {
    PCWSTR path = nullptr;
    BOOL handled = FALSE;
};

}  // namespace qttabbar::hooks

