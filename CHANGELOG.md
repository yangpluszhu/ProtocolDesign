# Changelog

All notable changes to this project will be documented in this file.

This project loosely follows the principles of [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/).

## [Unreleased]

### Planned

- Expand beyond the current retrospective cohort protocol module
- Continue improving UI polish and multi-resolution usability
- Improve model-provider compatibility and generation robustness

## [1.0.0] - 2026-05-24

### Added

- Initial public repository materials for GitHub, including README, LICENSE, CONTRIBUTING guide, and release notes
- Windows desktop workflow for generating Chinese retrospective cohort study protocol DOCX documents from study summaries
- Support for `.docx`, `.md`, and `.txt` study summary inputs
- Provider support for DeepSeek, Kimi, GLM, MiMo, and custom OpenAI-compatible endpoints
- DOCX rendering pipeline for cover page, table of contents, structured summary table, protocol sections, and appendix tables
- Run brief output as `_运行简报.txt`
- Optional local API key persistence encrypted with Windows DPAPI
- Build pipeline for packaging `ProtocolDesign.exe` into the `APP/` directory
- Help-document generation pipeline in Markdown, DOCX, and packaged release artifacts

### Changed

- Reworked the desktop UI into a grouped workbench layout better suited to scientific and academic usage
- Improved layout behavior across different desktop resolutions
- Updated help screenshots and synchronized help documentation with the current UI

### Fixed

- Improved robustness around provider-specific response handling and JSON repair paths
- Hardened parts of file-name handling, help opening, output-directory opening, and build-time help generation behavior
