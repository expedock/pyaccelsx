# Changelog

This is the changelog/release notes for `Pyaccelsx`.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [0.2.2] - 2024-09-10

### Added

- Make `write_string` automatically truncate string if it exceeds the max length allowed in Excel.

### Fixed

- Fixed outdated documentation in README
- Added missing changes for `row` and `column` types for `write_null`.

### Removed

- Removed misleading documentations.

## [0.2.1] - 2024-09-08

### Added

- Added formatting support for `bg_color`.
- Added a writer handler for `None` and `bool` values.
- Added class and function docs.
- Categorize writers into a single module.

### Changed

- Use one single function to write for multiple types (`str`, `int`, `float`, `bool`, and `None`).
- Use `write_string` and `write_string_with_format` to write strings.
- Use index when getting a worksheet instead of name.
- Make rows and columns to follow the correct type from `rust_xlsxwriter`.
- Updated README.

### Fixed

- Fixed `Format` object always getting created when writing numbers.
- Enabled `ExcelFormat` to be cloned.
- Removed redundant functions.

### Removed

- Removed unused files.

## [0.2.0] - 2024-09-04

### Added

- Added formatting support for the following format options: `bold`, `underline`, `alignment`, `font_color`, `border`, `num_format`.
- Added support for writing string and numbers.
- Added support for merging cell range.
- Added support for merging cells and write into the merged cells (numbers and strings).
- Added support for setting column width.
- Added support for freeze panes.

### Fixed

- Fixed the structure of files.

## [0.1.0] - 2024-09-04

### Added

- Initial version.
