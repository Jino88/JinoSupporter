# ReviewCase Batch Summary

Generated at: 2026-07-06T12:35:19Z

## Counts

- excluded: 7
- manual_draft_exists: 1
- needs_ai_verification: 922
- needs_review: 60
- total: 989

## Outputs

- Manifest: `REVIEWCASE_AI_DRAFTS\batch\reviewcase_batch_manifest.json`
- File drafts: `REVIEWCASE_AI_DRAFTS\batch\files`

## Notes

- This batch does not modify the MicroSpeaker SQLite database.
- User-confirmed manual drafts are not overwritten.
- `needs_ai_verification` drafts must be verified by AI/user before Ask AI treats them as final ReviewCases.
- Excluded files follow `REVIEWCASE_AI_AUDIT_DECISIONS.md` or lack citeable extracted row/cell evidence.
