# custom-reports

## Problem

Custom reports need to be generated based on key metrics in the GRC program.

Drata does not produce a formatted executive summary document out of the box. The data is in the platform but getting it into a deliverable format meant pulling numbers by hand every quarter and building the document manually. This script does that automatically.

## What It Does

Connects to the Drata Public API V1, pulls controls filtered to a specific framework, pulls all monitor test results, and computes four metric categories from the raw data:

- Control status: pass, fail, pending, derived from live monitor results
- Policy implementation: tiered from control flags (isReady, hasEvidence, isMonitored, hasOwner)
- Implementation fullness: how completely each control is implemented across systems
- Automation coverage: how much of the framework has automated testing behind it

Outputs a .docx with four sections (Executive Summary, Scope, Findings, Remediation) and four embedded pie charts. All numbers come from the API at runtime. The file is ready to send when the script finishes.

## Requirements

- Python 3.10 or later
- Dependencies: `pip install -r requirements.txt`

## Usage

### Basic

```
python report_generator.py --token TOKEN --company "Suncoast" --quarter 1 --year 2026
```

### Arguments

| Argument | Required | Default | Description |
|---|---|---|---|
| `--token` | Yes | | Drata API bearer token |
| `--company` | No | `Your Organization` | Company name used in the report narrative |
| `--framework` | No | `CIS 8` | Framework to report on, must match a supported tag (see below) |
| `--quarter` | Yes | | Report quarter: 1, 2, 3, or 4 |
| `--year` | No | Current year | Report year |
| `--cis-range` | No | Auto-detected from control codes | Control range label for the narrative, e.g. `1-18` |
| `--output` | No | `GRC_Summary_<Company>_Q<N>_<Year>.docx` | Output file path |
| `--insecure` | No | False | Disable SSL certificate verification, needed on some corporate networks |
| `--proxy` | No | | Proxy URL, e.g. `http://proxy.corp.example.com:8080` |

### Supported Frameworks

| `--framework` value | Notes |
|---|---|
| `CIS 8` | Default |
| `SOC 2` | |
| `ISO 27001` | |
| `ISO 27001:2022` | |
| `HIPAA` | |
| `PCI DSS` | |
| `PCI DSS v4` | |
| `GDPR` | |
| `CCPA` | |
| `NIST 800-53` | |
| `NIST CSF` | |
| `NIST CSF 2.0` | |
| `NIST 800-171` | |
| `CMMC` | |
| `FedRAMP` | |
| `HITRUST` | |
| `SOX ITGC` | |
| `DORA` | |
| `NIS2` | |
| `Cyber Essentials` | |
| `Essential Eight` | |

Full mapping in `FRAMEWORK_TAG_MAP` inside `report_generator.py`.

### API Token

Generate a token in Drata under Settings > API. Assign read access to Controls and Monitors. Pass it via `--token`.

### Example: Full Run

```
python report_generator.py ^
    --token 47b2c5f9-xxxx ^
    --company "Suncoast" ^
    --framework "CIS 8" ^
    --quarter 1 ^
    --year 2026 ^
    --cis-range "1-18" ^
    --output "Suncoast_GRC_Q1_2026.docx"
```
