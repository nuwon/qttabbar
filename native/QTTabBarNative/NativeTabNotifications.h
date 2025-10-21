#pragma once

#include <commctrl.h>

class TabItemData;

constexpr UINT NTCN_FIRST = 0U - 200U;
constexpr UINT NTCN_TABSELECT = NTCN_FIRST - 0;
constexpr UINT NTCN_TABCLOSE = NTCN_FIRST - 1;
constexpr UINT NTCN_TABCONTEXT = NTCN_FIRST - 2;
constexpr UINT NTCN_BEGIN_DRAG = NTCN_FIRST - 3;
constexpr UINT NTCN_END_DRAG = NTCN_FIRST - 4;
constexpr UINT NTCN_RELAY_TOOLTIP = NTCN_FIRST - 5;
constexpr UINT NTCN_DRAGGING = NTCN_FIRST - 6;

struct NMTABSTATECHANGE {
    NMHDR hdr;
    int index;
    TabItemData* tab;
};

struct NMTABDRAG {
    NMHDR hdr;
    int index;
    POINT screenPos;
};

