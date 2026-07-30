# Change Log

## 1.0.11 - 2026-05-25

### Standalone
- Re-synced all current Web-managed BMES/NG Rate files into the standalone app.
- Verified the Web-to-Standalone sync set has no mismatched files after namespace/rendermode conversion.
- Rebuilt the standalone update package so clients receive the latest Web changes.

## 1.0.10 - 2026-05-25

### Web
- Added explicit `Update To Server` actions to Routing Table and Reason Table.
- Stopped automatic Routing/Reason server sync after local add, edit, delete, import, or refresh.
- Stopped automatic server table pull when NG Rate processing starts; local Routing/Reason tables are used until the user updates the server.
- Kept Routing/Reason edits saved locally first, with status messages telling the user to use `Update To Server`.
- Removed visible mojibake from NG Rate view tabs and group headers.
- Added standalone update manifest and package endpoints for published standalone releases.

### Standalone
- Synced the current Web NG Rate, Routing Table, Reason Table, settings, and service changes into the standalone app.
- Added Routing Table and Reason Table `Update To Server` buttons.
- Changed first-run NG Rate storage paths so they start unset and must be configured by the user.
- Added standalone app update support and VS Code launch entries for standalone build/update workflows.
- Bumped standalone version to `1.0.10`.
