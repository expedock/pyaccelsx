# Changelog

This is the changelog/release notes for `Pyaccelsx`.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.0] - 2024-09-06

### Added

- Added formatting support for `bg_color`.
- Added a writer handler for `None` and `bool` values.
- Added class and function docs.

### Changed

- Use `write_string` and `write_string_with_format` to write strings.
- Use index when getting a worksheet instead of name.

### Fixed

- Fixed `Format` object always getting created when writing numbers.

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