# Changelog

All notable changes to this project are listed here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Changed

- License changed from MIT to the GNU General Public License v3.0 (GPL-3.0-only). Version 1.0.0 remains under MIT.

## [1.0.0] - 2026-10-21

First public release.

### Added

- Highlights the row and column of the active cell with a transparent layer above the grid; the workbook itself is
  never changed.
- Two modes on the **View** tab, group **Focus**: **Always**, and **On search** (only while *Find and Replace* is open).
- Right-to-left sheets, frozen panes, split windows, merged cells, zoom, 100% / 125% / 150% display scaling, several
  workbooks side by side.
- Terminal server / RDS support, including disconnect and reconnect.
- Central off switch for administrators (`admin\Set-FocusOverlayPolicy.ps1`).
- Install scripts for one user or for every user of a computer, and an uninstall script.
- Hebrew and English ribbon labels, chosen by Office's UI language.

[1.0.0]: https://github.com/3en5/excel-focus-cell/releases/tag/v1.0.0
