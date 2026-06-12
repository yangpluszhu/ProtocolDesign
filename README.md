# ProtocolDesign

> Clinical Research Protocol Assistant — An Automated Retrospective Cohort Study Protocol Generation Tool

**ProtocolDesign** is a Windows desktop application designed for Chinese medical research scenarios, developed by the [Clinical Research Center, Longhua Hospital, Shanghai University of Traditional Chinese Medicine](https://www.longhuahospital.com.cn). It automatically expands a brief **research summary** into a structurally complete, professionally formatted **Chinese retrospective cohort study protocol DOCX document**.

The software connects to external inference service providers (e.g., DeepSeek, Kimi, GLM, MiMo) and leverages built-in cohort-protocol skill rules and template writing guidelines to transform concise study summaries into publication-ready research protocols.

## 🎯 Key Benefits

| Traditional Workflow | With ProtocolDesign |
|---------------------|---------------------|
| Manual chapter-by-chapter writing, taking days | Assisted full-protocol generation in minutes |
| Inconsistent structure and formatting | Built-in 50+ standardized section headings and DOCX template |
| Missing key methodological elements | Automated validation of follow-up windows, outcome observation windows, and appendix completeness |
| Requires professional typesetting | One-click generation of cover page, table of contents, summary table, body text, and appendix tables |

## ✨ Features

- **📥 Multi-format Summary Input** — Supports `.docx`, `.md`, and `.txt` study summaries
- **🔌 Multi-platform Integration** — Built-in configurations for DeepSeek, Kimi, GLM, MiMo; supports any OpenAI-compatible API endpoint
- **🧠 Inference Mode Adaptation** — Automatically enables/disables reasoning mode per provider (maximum intensity for DeepSeek/GLM/MiMo; disabled per version requirements for Kimi)
- **📄 Professional DOCX Output** — Cover page, updatable table of contents, structured summary table, complete methodology sections, appendix tables
- **✔️ Automated Quality Validation** — Section completeness, follow-up time-window, outcome observation window, template artifact and placeholder detection
- **🔄 Multi-layer Fault Tolerance** — JSON repair (control characters / encoding errors / trailing commas), streaming decode, retry with backoff, Kimi segmented assembly
- **🔐 Secure Key Storage** — Optional encryption of API keys using Windows DPAPI per user identity
- **📋 Run Brief** — Automatically outputs a brief report including input parameters, service information, and risk warnings after each generation

## 📑 Generated Protocol Structure

The output DOCX document includes the following complete structure:

```
Cover Page (study title, center, investigator, version info, etc.)
Table of Contents
Protocol Summary (structured two-column table)
  1.  Background and Significance
  2.  Study Objectives (primary / secondary)
  3.  Study Design (type, time zero, bias minimization)
  4.  Study Subjects (data source, target population, inclusion/exclusion criteria, sample size)
  5.  Exposure Definition and Selection
  6.  Comparator Group Definition and Selection
  7.  Comparability of Indications, Contraindications, and Treatment Intensity
  8.  Concomitant Medications, Drug Switching, Discontinuation, Compliance, and Misclassification
  9.  Outcome Definition and Measurement (primary / secondary / safety / competing events)
 10.  Covariate Selection and Measurement
 11.  Statistical Analysis Methods (confounding control, subgroup, missing data, sensitivity analyses)
 12.  Data Management
 13.  Ethical Considerations
      Appendix: Recommended Attachments and Forms
      Appendix Table 1: Operational Definitions of Exposure, Comparator, Outcomes, and Covariates
      Appendix Table 2: Sensitivity Analysis Matrix
```

## 🔗 Supported Inference Service Providers

| Provider | Default Engine | Base URL | Reasoning Mode |
|----------|---------------|----------|---------------|
| **DeepSeek** | deepseek-v4-pro | `https://api.deepseek.com/v1` | ✅ Maximum intensity |
| **Kimi** | kimi-k2.6 | `https://api.moonshot.cn/v1` | ❌ Disabled per version (segmented generation) |
| **GLM** | glm-4.5 | `https://open.bigmodel.cn/api/paas/v4` | ✅ Maximum intensity |
| **MiMo** | mimo-v2.5-pro | `https://api.xiaomimimo.com/v1` | ✅ Maximum intensity |
| **Custom** | User-specified | User-specified | Compatibility mode |

> All providers communicate through the OpenAI-compatible `chat/completions` API. You must register an account with the respective platform and obtain an `api_key` before use.

## 🏗️ Architecture Overview

```
User Input (summary document + service configuration)
       │
       ▼
  ┌─ Rule Layer ─────────────────────────┐
  │  SKILL.md (cohort-protocol skill)     │
  │  template-writing-guide.md (template) │
  │  Dynamic prompt assembly               │
  └──────────────┬───────────────────────┘
                 │
                 ▼
  ┌─ Invocation Layer ────────────────────┐
  │  Per-provider independent workflows   │
  │  · Reasoning mode / JSON mode fallback│
  │  · Streaming / non-streaming switch   │
  │  · Transient overload auto-retry      │
  │    (429 / 5xx)                        │
  │  · Kimi cover-summary + section-batch│
  │    assembly                           │
  └──────────────┬───────────────────────┘
                 │
                 ▼
  ┌─ Validation Layer ────────────────────┐
  │  JSON parsing + control-char repair   │
  │  Section completeness validation       │
  │  Follow-up / outcome window checks    │
  │  Template artifact & placeholder scan  │
  │  Auto-retry on validation failure      │
  └──────────────┬───────────────────────┘
                 │
                 ▼
  ┌─ Rendering Layer ────────────────────┐
  │  protocol_renderer                    │
  │  · Applies corhortCRU-small.docx     │
  │    template                           │
  │  · Cover → TOC → Summary → Body →    │
  │    Appendix                            │
  │  · Automated font, color, and table   │
  │    formatting                         │
  └──────────────┬───────────────────────┘
                 │
                 ▼
    Output DOCX + Run Brief TXT
```

## 📁 Project Structure

```
ProtocolDesign/
├── src/
│   ├── ProtocolDesign.py            # Main application: GUI + service invocation + workflow orchestration
│   ├── protocol_renderer.py         # DOCX rendering engine (JSON → Word document)
│   ├── build_help_docx.py          # Help documentation DOCX generator
│   ├── resource_crypto.py           # Resource encryption/decryption (AES-256-GCM, dev/release dual mode)
│   └── encrypt_resources.py         # Build-time resource encryption tool
├── resources/
│   ├── SKILL.md                    # cohort-protocol skill rules (prompt core)
│   ├── references/
│   │   └── template-writing-guide.md  # Template writing guide (prompt core)
│   └── assets/
│       └── corhortCRU-small.docx   # DOCX formatting template
├── docs/
│   ├── ProtocolDesign_帮助文档.md
│   ├── ProtocolDesign_帮助文档.docx
│   └── ProtocolDesign_帮助文档.pdf
├── help_assets/                     # UI screenshots
├── samples/                         # Example summary documents
├── packaging/                       # PyInstaller packaging configuration
├── build_windows.ps1                # Windows build script
├── requirements.txt                 # Python dependencies
├── CHANGELOG.md                     # Version changelog
├── CONTRIBUTING.md                  # Contribution guidelines
├── LICENSE                          # GPL-3.0 license
└── README.md                        # This file
```

## 🚀 Getting Started

### Prerequisites

- Windows 10 / 11
- Python 3.11+
- One or more inference service provider accounts (DeepSeek / Kimi / GLM / MiMo or any OpenAI-compatible endpoint)

### Run from Source

```powershell
# 1. Clone the repository
git clone https://github.com/yangpluszhu/ProtocolDesign.git
cd ProtocolDesign

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch the application
python src/ProtocolDesign.py
```

### Build Executable

```powershell
# Full build (install deps + syntax check + help docs + encrypt resources + package EXE)
.\build_windows.ps1

# Skip dependency installation and PDF generation
.\build_windows.ps1 -SkipInstall -SkipPdf
```

The built executable will be located at `APP/ProtocolDesign.exe`.

## 📖 Usage

1. **Select Module** — Choose "Retrospective Cohort Study Protocol" (the only module available in V1.0)
2. **Select Summary Document** — Click browse and pick a prepared `.docx` / `.md` / `.txt` summary file
3. **Configure Service Endpoint** — Select a provider, then enter `base_url`, `api_key`, and the inference engine name
4. **Set Output Location** — Choose an output directory and filename
5. **Click Generate** — Wait for protocol expansion and DOCX rendering to complete
6. **Manual Review** — Open the generated document in Word, update the table of contents, and perform a line-by-line review

> 💡 **Summary Tips**: For best results, your summary should include the study title, data source, target population, inclusion/exclusion criteria, exposure and comparator definitions, outcome measures, statistical analysis considerations, and ethical requirements. The more specific the summary, the more reliable the output.

> ⚠️ For detailed instructions, see [`docs/ProtocolDesign_帮助文档.md`](docs/ProtocolDesign_帮助文档.md) or click the "Open Help Document" button in the application.

## 🛡️ Security

| Measure | Description |
|---------|-------------|
| Key Encryption | API keys are encrypted with Windows DPAPI bound to the current user identity; ciphertext can only be decrypted on the same machine by the same user |
| Isolated Storage | Keys for different providers are stored separately; unchecking the save option immediately deletes the corresponding entry |
| No Key Logging | API keys are never recorded in run briefs, error logs, or help documentation |
| Resource Encryption | Skill rules and templates are encrypted with AES-256-GCM at build time; plaintext is never included in the distributable package |

## ⚖️ Risk Disclaimer

> **Notice: This protocol is generated by a software tool and is intended solely as a reference for the research team. It should not be regarded as a final conclusion or decision basis.** Tool-generated content may contain factual inaccuracies, insufficient logical inferences, incomplete or outdated references, or biases in understanding specific research scenarios. It may also omit critical variables, ethical compliance requirements, data security considerations, or intellectual property risks. Before adoption, the research team should verify each element — study hypotheses, technical approach, data sources, study design, compliance requirements, and feasibility — using professional judgment, and confirm the protocol through review by relevant domain experts. Any research, submissions, or implementation based on this protocol shall be subject to manual review and formal validation. The software developer bears no responsibility for any adverse consequences arising from the use of documents generated by this software.

## 🛠️ Tech Stack

| Component | Technology |
|-----------|-----------|
| GUI Framework | Python tkinter |
| Document Generation | python-docx |
| Resource Encryption | AES-256-GCM (cryptography) |
| Key Storage | Windows DPAPI (ctypes) |
| HTTP Requests | requests |
| Image Processing | Pillow (help doc screenshot annotation) |
| Packaging & Distribution | PyInstaller (--onefile --windowed) |

## 🤝 Contributing

Contributions are welcome! Please refer to [`CONTRIBUTING.md`](CONTRIBUTING.md) for environment setup, development workflow, and pre-submission checklist.

## 📄 License

This project is open-sourced under the [GPL-3.0](LICENSE) license.

## 📬 Contact

- **Developer**: Clinical Research Center, Longhua Hospital, Shanghai University of Traditional Chinese Medicine
- **Email**: yangpluszhu@sina.com
- **GitHub**: [https://github.com/yangpluszhu/ProtocolDesign](https://github.com/yangpluszhu/ProtocolDesign)

---

*ProtocolDesign V1.0 © 2025 Clinical Research Center, Longhua Hospital, Shanghai University of Traditional Chinese Medicine*
