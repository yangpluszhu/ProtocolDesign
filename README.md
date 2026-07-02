# Clinical Research Protocol Assistant (ProtocolDesign)

> A Windows desktop tool for Chinese medical research scenarios. It automatically expands a research summary into a structurally complete, properly formatted retrospective cohort study protocol as a DOCX document.

The Clinical Research Protocol Assistant (ProtocolDesign) focuses on clinical research and the standardization of medical writing. The current version lets a research team import a protocol summary into the desktop application, generates a complete research protocol through an OpenAI-compatible large language model endpoint, and outputs a Word document that can be further reviewed and edited.

## Why use ProtocolDesign

- **Built for real research-writing workflows**: The output is a `.docx` document you can review and keep editing, not just plain text to copy.
- **Preserves the protocol structure**: Automatically generates the cover page, table of contents, structured summary table, every body chapter, and the appendix tables.
- **Multi-provider compatible**: Supports DeepSeek, Kimi, GLM, MiMo, and any custom OpenAI-compatible service.
- **Better suited for Chinese medical writing**: Ships with built-in section-writing rules, prompt resources, and protocol-template constraints.
- **Ready to ship as a desktop app**: Can be packaged as `ProtocolDesign.exe` for easy distribution to research teams on Windows.
- **API keys stored encrypted locally**: Optionally save keys encrypted with Windows DPAPI, never written to configuration in plain text.

## Current Capabilities

The current version officially supports only:

- `Retrospective Cohort Study Protocol` (`回顾性队列研究方案`)

Other modules shown in the UI are reserved for future upgrades and are not used for official generation in this version.

## Generation Output

On success, the software writes the following to the target directory:

- A complete research protocol `.docx`
- A run brief `_运行简报.txt` sharing the same base name

If generation fails, the program additionally writes to its own directory:

- `ProtocolDesign_error.log`

## Quick Start

### Requirements

- Windows 10 / 11
- Python 3.11+
- Access to a large language model service that exposes an OpenAI-compatible `chat/completions` endpoint
- LibreOffice (optional, used only to generate the help-document PDF)

### Install Dependencies

```powershell
python -m pip install -r requirements.txt
```

### Run from Source

```powershell
python src/ProtocolDesign.py
```

### Build a Release

```powershell
.\build_windows.ps1
```

After the build finishes, release artifacts are written to the `APP/` directory.

## Usage Workflow

1. Launch `ProtocolDesign.exe`, or run the source version.
2. Select a module — for now, choose `回顾性队列研究方案` (Retrospective Cohort Study Protocol).
3. Import the protocol summary file (`.docx`, `.md`, or `.txt` are supported).
4. Configure the model provider, `base_url`, `api_key`, and model name.
5. Choose the output directory and enter the output file name.
6. Click “确定生成方案” (Generate Protocol) and wait for the model generation and DOCX rendering to finish.

## Project Structure

```text
ProtocolDesignCode/
├─ src/
│  ├─ ProtocolDesign.py          # Main app: UI, model invocation, JSON repair, and generation workflow
│  ├─ protocol_renderer.py       # DOCX renderer
│  ├─ build_help_docx.py         # Help-document generator script
│  ├─ resource_crypto.py         # Decrypt/load resources after packaging
│  └─ encrypt_resources.py       # Build-time resource encryption
├─ resources/                    # Prompt resources, writing guides, template assets
├─ docs/                         # Markdown help documentation
├─ help_assets/                  # README / help-document screenshot assets
├─ samples/                      # Example summary files
├─ APP/                          # Packaging output directory
├─ build_windows.ps1             # One-click Windows packaging script
└─ requirements.txt
```

## Documentation

- User guide: [`docs/ProtocolDesign_帮助文档.md`](docs/ProtocolDesign_帮助文档.md)
- Contributing guide: [`CONTRIBUTING.md`](CONTRIBUTING.md)
- Release notes: [`CHANGELOG.md`](CHANGELOG.md)

## Development Notes

### Common Commands

```powershell
python -m py_compile src/ProtocolDesign.py src/protocol_renderer.py src/build_help_docx.py src/resource_crypto.py src/encrypt_resources.py
python src/build_help_docx.py
python src/encrypt_resources.py
```

### Implementation Overview

```text
Protocol summary file
  → Read text
  → Assemble prompt (skill rules + writing guide + user summary)
  → Call the LLM endpoint to generate JSON
  → Fix common JSON formatting issues
  → Validate structural completeness
  → Render into a DOCX research protocol
  → Write out the run brief
```

## Disclaimer

Output from this software is AI-assisted and is provided for the research team’s reference only. It must not be treated as a final research conclusion, an official submission text, or a basis for clinical or research decisions. Before use, please apply professional judgment and manually review the study design, variable definitions, statistical analysis, and ethical and data-security requirements.

## License

This project is released under the [`GPL-3.0`](LICENSE) license.

## Contact

- Author email: yangpluszhu@sina.com
- Author GitHub: [https://github.com/yangpluszhu](https://github.com/yangpluszhu)
