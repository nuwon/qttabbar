# Registry Validation Checklist

To verify the native configuration reader/writer matches the managed build:

1. Import `QTTabBar/Resources/old/QTTabBarSettings-2023-1-16.reg` into a clean test profile.
2. Launch the native bridge and call `LoadConfigFromRegistry()`, confirming that every category (`Window` through `Desktop`) populates without defaults.
3. Compare parsed data against the `.reg` file for spot-check keys:
   - `Config\Tabs`: enum values (`NewTabPosition`, `NextAfterClosed`) should respect their string forms (e.g., `"Rightmost"`, `"LastActive"`).
   - `Config\Skin`: colors (`TabTextActiveColor`, `RebarColor`) must deserialize from the managed JSON objects; image paths should be cleared if invalid.
   - `Config\Keys`: ensure `Shortcuts` array length matches `BindAction.KEYBOARD_ACTION_COUNT` and plugin dictionaries omit null entries.
4. Modify several settings in memory (e.g., tweak `TabHeight`, remove a plugin button, change preview limits) and persist via `WriteConfigToRegistry`.
5. Export the registry branch and diff against the sample export to confirm formatting: JSON payloads, string enums, and DWORD flags mirror managed expectations.
6. Re-load to ensure `ApplyConfigValidation` performs clamping (`Tips.PreviewMaxWidth`, `Skin.TabMinWidth`, etc.) and path verification (`Skin.TabImageFile`, `BBar.ImageStripPath`).

These steps mirror the managed configuration pipeline and provide manual coverage for the native registry serializers.
