# Security policy

## Supported versions

Only the latest release receives fixes.

## Reporting a vulnerability

Please **do not** open a public issue for a security problem. Report it privately through GitHub instead:
**Security** tab → **Report a vulnerability** on this repository.

Include the Excel and Windows versions, what you did, and what happened. You will get a reply within a few working
days.

## Scope

Excel Focus Cell runs inside Excel with the user's own rights. It has no network access, reads only positions from
Excel's object model (never cell values), and writes only to its own registry key and a local log file.
