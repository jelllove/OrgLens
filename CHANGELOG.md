# Changelog

## 0.2.2

- All four color selectors show a color swatch and name in both the dropdown
  and the selected field, using Outlook's 17 supported palette values.
- The settings window is modeless: keep working in Outlook while it stays open.
  Clicking the Ribbon button again activates the existing window, including
  restoring it from minimized state.
- Normal Close keeps the unsaved-change confirmation. Outlook shutdown or
  add-in disconnection disposes the window without leaving a detached editor.
  Save your portable settings before exiting Outlook.
- Adds color-rendering and modeless lifecycle regressions to the UI smoke suite.
- Publishes the source under MIT, with a screenshot in the repository README,
  Git ignore rules, a distributable package, and a SHA-256 checksum.

Settings schema version 2 and all existing native rule behavior are unchanged.
Close Outlook before installing this update. Existing JSON settings remain compatible.

## 0.2.1

- Fixes Outlook clearing formatting conditions after `View.Save()` or `View.Apply()`.
  Saves `AutoFormatRules`, activates through `Explorer.CurrentView`, and verifies
  the persisted filters before reporting success.
- Disables OrgLens rules on failed condition verification and restores native
  rule/font snapshots on rollback without re-enabling old empty conditions.

Users upgrading from 0.2.0 must Refresh and Apply again after installation;
updating the add-in alone does not repair previously saved mailbox views.

## 0.2.0

- Groups styles into BOSS read/unread, custom unread senders, and explicit To/Cc.
- Adds font size, live previews, and portable validated JSON Open/Save/Save As.

## 0.1.0

- Initial classic Outlook COM add-in and fictional standalone preview.
- Exchange address-book manager-chain discovery and native view formatting.
