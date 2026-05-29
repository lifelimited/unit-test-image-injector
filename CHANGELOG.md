# Changelog

All notable changes to this project will be documented in this file.

## [1.3.2] - 2026-05-30
### Added
- **Universal Document Support**: The application can now process any document image (not just IP-phones), intelligently calculating orientation based on OCR text density.
- **Improved Smart-Rotation Algorithm**: Replaced keyword-specific tracking with an adaptive text-density scoring algorithm to robustly handle generic documentation processing.

### Fixed
- **HEIC Processing Bug**: Resolved a bug where `.HEIC` and `.HEIF` images were incorrectly bypassed during the format conversion phase, ensuring reliable processing and injection.

## [1.2.1] - Previous Versions
- Added support for macOS compatibility (directory separator normalization).
- Initial framework for image processing and document injection based on IP-Phone specific tracking (keyword boosting).
