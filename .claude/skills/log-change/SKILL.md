---
name: log-change
description: Enforces the CLAUDE.md change log policy — update the change log table before marking any change as complete
user-invocable: false
---

Before marking any change as complete, update the change log table in CLAUDE.md:

- Add new entries with status `PROPOSED` when a change is first discussed
- Update to `DONE` only after the change is implemented and confirmed
- Update to `REJECTED` if the change is abandoned or refused — always include a reason in the Notes column
- Add entries to the correct version section (currently v0.3.1 or v0.3.2)
- Use the exact table format: `| # | Change | Status | Notes |`
- Number entries sequentially within their version block
- Never silently apply or drop a change without a log entry

This applies to every change, no matter how small.
