# Loading Timestamp — One-table loading report

## V16: Single direct loading report

The app reads/writes only `LOADING_REPORT` for vehicle data. Every vehicle row contains its own date, shift, incharge, helpers and times. No `ARRAYFORMULA`, header lookup, or separate item table is used after cutover. A:I retain the original report column positions; J:O add SHIFT, HELPER COUNT, ENTRY_ID, SHIFT_ID, CHECK, UPDATED_AT. IDs are for stable edit/delete, not a lookup in another table.

### Activate on the existing Google Sheet

1. Replace only `LoadingWebAppConnector.gs` with this repository's version. Keep existing `Code.gs` and other scripts.
2. Save, then Deploy → Manage deployments → Edit → New version → Deploy, retaining existing access settings and URL.
3. Select `setupLoadingDirectReport` in the editor and Run once. Authorize the requested Sheet/trigger access. This imports the latest legacy records by existing item IDs, excludes tombstoned deletions, preserves a report backup, and hides LOADING_HEADER/LOADING_ITEMS as archives. Re-running after successful setup does not reimport old data.
4. Reopen/refresh the mobile app. New entries, edits, and deletions now go only to LOADING_REPORT. Old AppSheet clients must not be used for new entries after cutover.

`TOTAL_HRS` is calculated by Apps Script in decimal hours. Crossing midnight is supported. `CHECK` flags durations over eight hours and start times outside broad shift windows for review; it never guesses or silently changes a date, shift, or AM/PM choice. The installed edit trigger recalculates a supervisor's manual row edits and flags missing/invalid fields. Manual metadata edits detach the row from its old UI shift group. Do not edit ENTRY_ID or SHIFT_ID.

The report conversion alone does not deploy Apps Script. Until the deployment/setup steps are completed, the old web app continues to use the legacy tables. Setup reimports the latest legacy data at the cutover, covering entries added during this interval. Existing stage-time processing and customer master remain unchanged.
