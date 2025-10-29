# TabBarHost native command verification

The following manual checklist covers the newly implemented native command bindings. These steps mirror the managed regression suite and should be executed on a Windows host with Explorer integration enabled.

## Prerequisites
- Install the built explorer band and button bar.
- Enable the tab bar, button bar, and SubDirTip options in QTTabBar.
- Prepare a test folder tree containing at least one nested directory, shortcut (*.lnk*) to a folder, and a small file for hashing.

## Bindings under test
- **Tab context menus**: Assign keyboard shortcuts to `ShowTabMenuCurrent` and `ShowToolbarMenu`. Trigger each shortcut and confirm the respective menus appear anchored to the active tab/toolbar. Ensure the tab menu appears even without a pre-selected tab.
- **SubDirTip execution**: Bind `ShowTabSubfolderMenu` and press the shortcut while a tab is focused. Verify the first subdirectory opens immediately (matching the managed `PerformClickByKey` behaviour).
- **Search focus**: Bind `FocusSearchBarReal` and `FocusSearchBarBBar`. Run each shortcut to confirm focus switches to the Explorer search box and the button-bar search box respectively.
- **File tools**: Bind `CopySelectedPaths`, `CopySelectedNames`, and `ChecksumSelected`. With one or more files selected, run each command and validate clipboard content or checksum dialog results. Re-run `ChecksumSelected` with no selection to confirm the dialog still opens with the "No files selected" notice.
- **Window controls**: Bind `ToggleTopMost`, `TransparencyPlus`, `TransparencyMinus`, and `SendToTray`. Execute in sequence and verify the Explorer window toggles top-most state, opacity adjusts in 12-point steps, and tray integration captures the active tab list.
- **Merge windows**: Open two Explorer instances, populate each with unique tabs (including locked tabs), and trigger `MergeWindows`. Confirm all tabs move into the active window with their locked states intact and the secondary window closes.
- **New item actions**: Bind `NewFolder` and `NewFile`. Execute each shortcut and ensure the expected item is created in the current directory and selected for rename.
- **Item-level verbs**: Bind `ItemOpenInNewTab`, `ItemOpenInNewWindow`, `ItemCut`, `ItemCopy`, `ItemDelete`, and `ChecksumItem`. Select a folder shortcut (*.lnk*) and verify `ItemOpenInNewTab`/`ItemOpenInNewWindow` resolve the shortcut. Execute the clipboard verbs and checksum command to ensure the shell verbs run without errors.

Record pass/fail results for each row. Any failure should include repro steps, logs (from the DebugView output), and screenshots when applicable.
