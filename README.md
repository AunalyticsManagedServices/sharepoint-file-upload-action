# SharePoint File Upload GitHub Action

> 🚀 Automatically sync files from GitHub to SharePoint with intelligent change detection and Markdown-to-HTML conversion

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Version](https://img.shields.io/badge/Version-3.1.0-blue)](https://github.com/AunalyticsManagedServices/sharepoint-file-upload-action)

## 📋 Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [Quick Start](#quick-start)
- [Configuration](#configuration)
  - [Required Parameters](#required-parameters)
  - [Optional Parameters](#optional-parameters)
  - [File Glob Patterns](#file-glob-patterns)
- [Usage Examples](#usage-examples)
- [Advanced Features](#advanced-features)
  - [Smart Sync](#smart-sync)
  - [Markdown Conversion](#markdown-conversion)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## Overview

This GitHub Action provides enterprise-grade file synchronization between GitHub repositories and SharePoint document libraries. It intelligently uploads only changed files, converts Markdown documentation to SharePoint-friendly HTML, and provides detailed sync statistics.

### Why Use This Action?

- **📁 Automated Documentation Sync**: Keep your SharePoint documentation in sync with your GitHub repository
- **⚡ Efficient Updates**: Only uploads new or modified files, saving time and bandwidth
- **📝 Markdown Support**: Automatically converts `.md` files to styled HTML with Mermaid diagram support
- **🔄 Reliable**: Includes retry logic and comprehensive error handling
- **📊 Detailed Reporting**: Shows exactly what was uploaded, skipped, or failed

## Key Features

| Feature | Description |
|---------|-------------|
| **Smart Sync** | Uses xxHash128 content hashing (or file size) to skip unchanged files |
| **Markdown → HTML** | Converts `.md` files to beautifully styled HTML for SharePoint viewing |
| **Mermaid Diagrams** | Renders Mermaid flowcharts/diagrams as embedded SVG |
| **Batch Upload** | Handles multiple files and maintains folder structure |
| **Large File Support** | Automatically uses chunked upload for files >4MB |
| **Special Character Handling** | Sanitizes filenames for SharePoint compatibility |

## Quick Start

### 1. Set Up SharePoint App Registration

Your SharePoint administrator needs to:
1. Register an app in Azure AD
2. Grant permissions:
   - `Sites.ReadWrite.All` (minimum for file operations)
   - `Sites.Manage.All` (optional, enables automatic FileHash column creation)
3. Provide you with:
   - Tenant ID
   - Client ID
   - Client Secret

### 2. Add Secrets to GitHub

Navigate to **Settings → Security → Secrets and variables → Actions** and add:

```
SHAREPOINT_TENANT_ID
SHAREPOINT_CLIENT_ID
SHAREPOINT_CLIENT_SECRET
```

### 3. Create Workflow

Create `.github/workflows/sharepoint-sync.yml`:

```yaml
name: Sync to SharePoint
on:
  push:
    branches: [main]
  workflow_dispatch:

jobs:
  sync:
    runs-on: ubuntu-latest
    steps:
      - name: Github Checkout/Clone Repo # This is required
        uses: actions/checkout@v4 # This is required

      - name: Upload to SharePoint
        uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
        with:
          file_path: "**/*"
          host_name: "yourcompany.sharepoint.com"
          site_name: "YourSite"
          upload_path: "Shared Documents/GitHub Sync"
          tenant_id: ${{ secrets.SHAREPOINT_TENANT_ID }}
          client_id: ${{ secrets.SHAREPOINT_CLIENT_ID }}
          client_secret: ${{ secrets.SHAREPOINT_CLIENT_SECRET }}
```

## Configuration

### Required Parameters

| Parameter | Description | Example |
|-----------|-------------|---------|
| `file_path` | Files to upload (supports glob patterns) | `"**/*"` |
| `host_name` | Your SharePoint domain | `"company.sharepoint.com"` |
| `site_name` | SharePoint site name | `"TeamDocs"` |
| `upload_path` | Target folder in SharePoint | `"Shared Documents/Reports"` |
| `tenant_id` | Azure AD Tenant ID | `${{ secrets.SHAREPOINT_TENANT_ID }}` |
| `client_id` | App registration Client ID | `${{ secrets.SHAREPOINT_CLIENT_ID }}` |
| `client_secret` | App registration Client Secret | `${{ secrets.SHAREPOINT_CLIENT_SECRET }}` |

### Optional Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `file_path_recursive_match` | `false` | Enable recursive glob matching |
| `max_retries` | `3` | Number of upload retry attempts |
| `force_upload` | `false` | Skip change detection, upload all files |
| `convert_md_to_html` | `true` | Convert Markdown files to HTML |
| `login_endpoint` | `"login.microsoftonline.com"` | Azure AD login endpoint |
| `graph_endpoint` | `"graph.microsoft.com"` | Microsoft Graph API endpoint |

### File Glob Patterns

Understanding glob patterns helps you select exactly which files to upload:

| Pattern | Description | Example Matches |
|---------|-------------|-----------------|
| `*.pdf` | All PDF files in root directory | `report.pdf`, `guide.pdf` |
| `**/*.pdf` | All PDF files in any directory | `docs/report.pdf`, `archive/2024/guide.pdf` |
| `docs/*` | All files directly in docs folder | `docs/readme.md`, `docs/config.json` |
| `docs/**/*` | All files in docs and subfolders | `docs/api/spec.yaml`, `docs/guides/intro.md` |
| `**/*.{md,txt}` | All markdown and text files | `readme.md`, `notes.txt`, `docs/guide.md` |
| `!**/*.test.*` | Exclude test files | Excludes `file.test.js`, `spec.test.md` |

#### Glob Pattern Tips

- **`*`** matches any characters within a single directory level
- **`**`** matches zero or more directory levels
- **`{a,b}`** matches either pattern a or b
- **`!pattern`** excludes files matching the pattern
- Always use forward slashes (`/`) even on Windows

## Usage Examples

### Example 1: Upload All Files and Folders 

```yaml
- name: Github Checkout/Clone Repo # This is required
  uses: actions/checkout@v4 # This is required

- name: Sync All Files and Folders
  uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    file_path: "**/*"
    file_path_recursive_match: true
    host_name: "company.sharepoint.com"
    site_name: "RepoDocumentation"
    upload_path: "Technical Docs/Latest"
    tenant_id: ${{ secrets.SHAREPOINT_TENANT_ID }}
    client_id: ${{ secrets.SHAREPOINT_CLIENT_ID }}
    client_secret: ${{ secrets.SHAREPOINT_CLIENT_SECRET }}
```

### Example 2: Upload Only Changed PDF Reports

```yaml
- name: Github Checkout/Clone Repo # This is required
  uses: actions/checkout@v4 # This is required

- name: Upload Changed PDF Reports
  uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    file_path: "reports/*.pdf"
    force_upload: false  # Smart sync enabled
    convert_md_to_html: false  # Keep markdown as-is
    host_name: "company.sharepoint.com"
    site_name: "Analytics"
    upload_path: "Monthly Reports"
    tenant_id: ${{ secrets.SHAREPOINT_TENANT_ID }}
    client_id: ${{ secrets.SHAREPOINT_CLIENT_ID }}
    client_secret: ${{ secrets.SHAREPOINT_CLIENT_SECRET }}
```

### Example 3: Convert and Upload Markdown Documentation

```yaml
- name: Github Checkout/Clone Repo # This is required
  uses: actions/checkout@v4 # This is required

- name: Convert MD to HTML and Upload
  uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    file_path: "**/*.md"
    file_path_recursive_match: true
    convert_md_to_html: true  # Converts .md to styled HTML
    host_name: "company.sharepoint.com"
    site_name: "KnowledgeBase"
    upload_path: "Articles"
    tenant_id: ${{ secrets.SHAREPOINT_TENANT_ID }}
    client_id: ${{ secrets.SHAREPOINT_CLIENT_ID }}
    client_secret: ${{ secrets.SHAREPOINT_CLIENT_SECRET }}
```

### Example 4: Government Cloud Deployment

```yaml
- name: Github Checkout/Clone Repo # This is required
  uses: actions/checkout@v4 # This is required

- name: Upload to GovCloud SharePoint
  uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    file_path: "compliance/**/*"
    host_name: "agency.sharepoint.us"
    site_name: "Compliance"
    upload_path: "FY2024"
    tenant_id: ${{ secrets.SHAREPOINT_TENANT_ID }}
    client_id: ${{ secrets.SHAREPOINT_CLIENT_ID }}
    client_secret: ${{ secrets.SHAREPOINT_CLIENT_SECRET }}
    login_endpoint: "login.microsoftonline.us"
    graph_endpoint: "graph.microsoft.us"
```

## Advanced Features

### Smart Sync with Content Hashing

When `force_upload` is `false` (default), the action:

1. **Checks existing files** in SharePoint before uploading
2. **Calculates xxHash128** checksums for ultra-fast content comparison
3. **Falls back to file size** comparison if hash metadata unavailable
4. **Skips unchanged files** to save time and bandwidth
5. **Reports statistics** showing files uploaded vs. skipped

**Note**: The action automatically creates a `FileHash` column in SharePoint (if permissions allow) to store content hashes for future comparisons. This provides more reliable change detection than timestamps, especially in CI/CD environments.

**Console Output Example:**
```
[OK] Smart sync mode enabled - unchanged files will be skipped
[✓] FileHash column is available for hash-based comparison
[?] Checking if file exists in SharePoint: README.md
[#] Local hash: a3f5c892... for README.md
[#] Remote hash: a3f5c892... for README.md
[=] File unchanged (hash match): README.md
[*] File changed (hash mismatch): config.json
[+] New file to upload: changelog.md

[STATS] Sync Statistics:
   - New files uploaded:         3
   - Files updated:              2
   - Files skipped (unchanged):  45
   - Total files processed:      50

[DATA] Transfer Summary:
   - Data uploaded:   15.3 MB
   - Data skipped:    125.7 MB
   - Total savings:   125.7 MB (45 files not re-uploaded)

[EFFICIENCY] 90.0% of files were already up-to-date
```

### Markdown Conversion

When `convert_md_to_html` is `true` (default):

1. **Converts Markdown to HTML** with GitHub-flavored styling
2. **Renders Mermaid diagrams** as embedded SVG images
3. **Uploads HTML instead of MD** for better SharePoint viewing
4. **Preserves formatting** including tables, code blocks, and task lists

#### Supported Markdown Features

- **Headers** (H1-H6) with automatic anchors
- **Tables** with alternating row colors
- **Code blocks** with syntax highlighting
- **Task lists** with checkboxes
- **Blockquotes** with visual indicators
- **Links** and **images**
- **Mermaid diagrams** (flowcharts, sequence diagrams, etc.)

#### Mermaid Diagram Example

```markdown
\```mermaid
graph TD
    A[Start] --> B{Decision}
    B -->|Yes| C[Process]
    B -->|No| D[End]
    C --> D
\```
```

This will be converted to an embedded SVG diagram in the HTML output.

### Filename Sanitization

The action automatically handles SharePoint's filename restrictions:

| Invalid Character | Replacement |
|-------------------|-------------|
| `#` | `＃` (fullwidth) |
| `%` | `％` (fullwidth) |
| `&` | `＆` (fullwidth) |
| `:` | `：` (fullwidth) |
| `<` `>` | `＜` `＞` (fullwidth) |
| `?` | `？` (fullwidth) |
| `\|` | `｜` (fullwidth) |

Reserved names (CON, PRN, AUX, etc.) are automatically prefixed with underscore.

## Troubleshooting

### Common Issues

#### 1. Authentication Failed

**Error:** `Failed to connect to SharePoint: (401) Unauthorized`

**Solution:**
- Verify your Client ID and Client Secret are correct
- Ensure the app registration has `Sites.ReadWrite.All` permissions
- Check that the Tenant ID matches your SharePoint instance

#### 2. File Not Found

**Error:** `No files or directories matched pattern`

**Solution:**
- Check your glob pattern syntax
- Enable `file_path_recursive_match: true` for nested directories
- Ensure files exist in the repository (use `ls` in a previous step to debug)

#### 3. Upload Timeout

**Error:** `The operation timed out`

**Solution:**
- Large files may take time; the action automatically retries
- Consider splitting large uploads into multiple actions
- Check your network connectivity

#### 4. Markdown Conversion Failed

**Error:** `Mermaid conversion failed`

**Solution:**
- Verify Mermaid syntax is correct
- Simple diagrams work best for SVG conversion
- Complex diagrams might need simplification

### Debug Mode

To enable verbose logging, add a step before the upload:

```yaml
- name: List files to upload
  run: |
    echo "Files matching pattern:"
    ls -la docs/**/*.md
```

## Performance Optimization

### Tips for Faster Syncs

1. **Use specific glob patterns** instead of `**/*`
2. **Enable smart sync** (default) to skip unchanged files
3. **Upload frequently** to minimize changes per sync
4. **Organize files logically** to use targeted patterns
5. **Schedule during off-peak hours** for large syncs

**Performance Note**: The action uses xxHash128 for content verification, which processes files at 3-6 GB/s on modern CPUs. This is 10-20x faster than traditional SHA-256 hashing, ensuring minimal overhead even for large files.

### Benchmarks

| Files | Size | Smart Sync | Force Upload |
|-------|------|------------|--------------|
| 10 files | 50 MB | ~15 seconds | ~45 seconds |
| 100 files | 500 MB | ~30 seconds | ~5 minutes |
| 1000 files | 5 GB | ~2 minutes | ~30 minutes |

*Note: Times vary based on network speed and file changes*

## Security Considerations

- **Never commit secrets** to your repository
- **Use GitHub Secrets** for all sensitive values
- **Limit permissions** to only required SharePoint sites
- **Review app permissions** regularly
- **Use branch protection** to control who can trigger uploads

## Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

### Development Setup

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Testing Locally

```bash
# Build the Docker image
docker build -t sharepoint-upload .

# Run with test parameters
docker run sharepoint-upload \
  "test-site" \
  "test.sharepoint.com" \
  "tenant-id" \
  "client-id" \
  "client-secret" \
  "test-path" \
  "*.md"
```

## Support

- **📝 Issues**: [GitHub Issues](https://github.com/AunalyticsManagedServices/sharepoint-file-upload-action/issues)
- **💬 Discussions**: [GitHub Discussions](https://github.com/AunalyticsManagedServices/sharepoint-file-upload-action/discussions)
- **📧 Contact**: [Support Email](mailto:support@aunalytics.com)

## Acknowledgments

- Built with [Office365-REST-Python-Client](https://github.com/vgrem/Office365-REST-Python-Client)
- Markdown parsing by [Mistune](https://github.com/lepture/mistune)
- Mermaid diagrams by [Mermaid-CLI](https://github.com/mermaid-js/mermaid-cli)

---

<div align="center">
Made with ❤️ by <a href="https://github.com/AunalyticsManagedServices">Aunalytics Managed Services</a>
</div>