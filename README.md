# SharePoint File Upload GitHub Action

> 🚀 Automatically sync files from GitHub to SharePoint with intelligent change detection and Markdown-to-HTML conversion

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Version](https://img.shields.io/badge/Version-3.1.0-blue)](https://github.com/AunalyticsManagedServices/sharepoint-file-upload-action)

## 📋 Quick Navigation

| Getting Started | Configuration | Features | Resources |
|:----------------|:--------------|:---------|:----------|
| [Overview](#-overview) | [Required Parameters](#required-parameters) | [Smart Sync](#smart-sync-with-content-hashing) | [Troubleshooting](#-troubleshooting) |
| [Quick Start](#-quick-start) | [Optional Parameters](#optional-parameters) | [Markdown Conversion](#markdown-conversion) | [Performance](#-performance) |
| [Usage Examples](#-usage-examples) | [Glob Patterns](#file-glob-patterns) | [Sync Deletion](#sync-deletion) | [Security](#-security) |
| | [Exclusion Patterns](#exclusion-patterns) | [Filename Sanitization](#filename-sanitization) | [Contributing](#-contributing) |

## 🎯 Overview

Seamlessly synchronize files from your GitHub repository to SharePoint document libraries. This action intelligently uploads only changed files, converts Markdown to SharePoint-friendly HTML, and maintains perfect sync between your repository and SharePoint.

### Why Use This Action?

| Benefit | Description |
|---------|-------------|
| 📁 **Automated Sync** | Keep SharePoint documentation current with your GitHub repository |
| ⚡ **Smart Uploads** | Only uploads new or modified files (typically skips 60-90% of files) |
| 📝 **Markdown Support** | Converts `.md` files to styled HTML with Mermaid diagram rendering |
| 🔄 **Bidirectional Sync** | Optional deletion of SharePoint files removed from repository |
| 📊 **Detailed Reports** | Clear statistics on uploads, skips, and failures |
| 🔒 **Enterprise Ready** | Supports GovCloud, Sites.Selected permissions, and large files |

## ✨ Key Features

| Feature | Description |
|---------|-------------|
| **Smart Sync** | xxHash128 content comparison skips unchanged files automatically |
| **Markdown → HTML** | GitHub-flavored HTML with embedded Mermaid diagrams |
| **Large Files** | Chunked upload for files >4MB with automatic retry logic |
| **Folder Structure** | Maintains directory hierarchy from repository to SharePoint |
| **Special Characters** | Auto-sanitizes filenames for SharePoint compatibility |
| **Multi-Cloud** | Commercial, GovCloud, and sovereign cloud support |

## 🚀 Quick Start

### 1. Get SharePoint Credentials

Your SharePoint administrator provides:
- **Tenant ID** (Azure AD tenant identifier)
- **Client ID** (Application ID from app registration)
- **Client Secret** (App password/secret value)

**Permissions needed** (choose one):
- ✅ **Recommended**: `Sites.Selected` + grant access to specific sites ([setup guide](#sitesselected-setup))
- Alternative: `Sites.ReadWrite.All` (tenant-wide access)
- Optional: `Sites.Manage.All` (enables automatic FileHash column creation)

### 2. Add GitHub Secrets

**Settings** → **Security** → **Secrets and variables** → **Actions** → **New repository secret**

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
      - name: Checkout Repository
        uses: actions/checkout@v4

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

That's it! Push to `main` or click **Run workflow** to sync.

## ⚙️ Configuration

### Required Parameters

| Parameter | Description | Example |
|-----------|-------------|---------|
| `file_path` | Files to sync (glob patterns supported) | `"**/*"`, `"docs/**/*.md"` |
| `host_name` | SharePoint domain | `"company.sharepoint.com"` |
| `site_name` | SharePoint site name | `"TeamDocs"` |
| `upload_path` | Target folder path | `"Shared Documents/Reports"` |
| `tenant_id` | Azure AD Tenant ID | `${{ secrets.SHAREPOINT_TENANT_ID }}` |
| `client_id` | App Client ID | `${{ secrets.SHAREPOINT_CLIENT_ID }}` |
| `client_secret` | App Client Secret | `${{ secrets.SHAREPOINT_CLIENT_SECRET }}` |

<details>
<summary><strong>📖 Detailed Parameter Guide</strong></summary>

#### `file_path` - Which files to upload

Supports glob patterns for flexible file selection:
- `"docs/readme.md"` - Single specific file
- `"*.pdf"` - All PDFs in root directory
- `"**/*.md"` - All Markdown files in all subdirectories
- `"reports/**/*"` - Everything in reports folder

See [File Glob Patterns](#file-glob-patterns) for comprehensive pattern guide.

#### `host_name` - Your SharePoint domain

The domain part of your SharePoint URL (without `https://`):
- Commercial: `"company.sharepoint.com"`
- GovCloud: `"agency.sharepoint.us"`
- Custom: `"sharepoint.customdomain.com"`

**How to find**: Open SharePoint in browser, copy domain from URL bar.

#### `site_name` - SharePoint site name

The site collection name from your SharePoint URL:
- URL: `https://company.sharepoint.com/sites/TeamSite`
- Site name: `"TeamSite"`

**Note**: Case-sensitive in some configurations.

#### `upload_path` - Target folder in SharePoint

Format: `"DocumentLibrary/FolderPath/Subfolder"`

Common libraries:
- `"Shared Documents"` - Default library
- `"Documents"` - Alternate default
- `"Site Assets"` - Web resources

Examples:
- `"Shared Documents"` - Library root
- `"Documents/Reports"` - Reports subfolder
- `"Shared Documents/Q1/Financial"` - Multi-level path

**Folders are created automatically if they don't exist.**

#### `tenant_id`, `client_id`, `client_secret` - Azure credentials

**How to find**:
1. [Azure Portal](https://portal.azure.com) → **Azure Active Directory**
2. **App registrations** → Select your app
3. Copy **Tenant ID** and **Application (client) ID**
4. **Certificates & secrets** → Create secret → Copy **Value** immediately

⚠️ **Store all three in GitHub Secrets** - never commit to repository.

</details>

### Optional Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `file_path_recursive_match` | `false` | Enable recursive directory traversal |
| `max_retries` | `3` | Upload retry attempts (1-10) |
| `force_upload` | `false` | Skip change detection, upload all files |
| `convert_md_to_html` | `true` | Convert Markdown to HTML |
| `exclude_patterns` | `""` | Comma-separated exclusion patterns |
| `sync_delete` | `false` | Delete SharePoint files not in repository |
| `sync_delete_whatif` | `true` | Preview deletions without deleting |
| `max_upload_workers` | `4` | Concurrent upload workers (1-10) |
| `max_hash_workers` | `CPU count` | Hash calculation workers (auto-detect) |
| `login_endpoint` | `"login.microsoftonline.com"` | Azure AD endpoint |
| `graph_endpoint` | `"graph.microsoft.com"` | Microsoft Graph endpoint |

<details>
<summary><strong>📖 Detailed Optional Parameters</strong></summary>

#### `file_path_recursive_match` - Recursive traversal

Enable to include all subdirectories:

```yaml
# Without recursive (default):
file_path: "docs/*"        # Only files directly in docs/

# With recursive:
file_path: "docs/**/*"
file_path_recursive_match: true  # Includes all subdirectories
```

#### `max_retries` - Retry attempts

- **Default**: 3 attempts
- **When to adjust**:
  - Unstable networks: Increase to 5
  - Fast failure: Decrease to 1
- Uses exponential backoff (2s, 4s, 8s...)

```yaml
max_retries: 5    # More resilient
max_retries: 1    # Fail fast
```

#### `force_upload` - Force all uploads

- **Default**: `false` (smart sync - only changed files)
- **Use `true` when**:
  - First-time sync
  - Troubleshooting sync issues
  - Intentionally re-uploading everything

⚠️ **Warning**: Uploads ALL files regardless of changes (slower).

💡 **Tip**: Smart sync typically skips 60-90% of files.

#### `convert_md_to_html` - Markdown conversion

- **Default**: `true` (auto-convert `.md` files)
- **Output**: Styled HTML with GitHub formatting + Mermaid diagrams
- **Use `false`** when you want raw `.md` files

```yaml
convert_md_to_html: true   # README.md → README.html (styled)
convert_md_to_html: false  # README.md → README.md (as-is)
```

#### `exclude_patterns` - File exclusions

Comma-separated list to skip files:

```yaml
# Python projects
exclude_patterns: "*.pyc,__pycache__,.pytest_cache"

# Node.js projects
exclude_patterns: "node_modules,*.log,dist,build"

# Sensitive data
exclude_patterns: ".env,.env.local,secrets.json,*.key"
```

See [Exclusion Patterns](#exclusion-patterns) for detailed guide.

#### `sync_delete` - Bidirectional sync

- **Default**: `false` (safer)
- **Use `true`** to delete SharePoint files removed from repository
- ⚠️ **Always test with `sync_delete_whatif: true` first!**

```yaml
sync_delete: false  # No deletions (default)
sync_delete: true   # Enable deletion (test with whatif first!)
```

See [Sync Deletion](#sync-deletion) for complete guide.

#### `sync_delete_whatif` - Preview deletions

- **Default**: `true` (preview mode)
- **Requires**: `sync_delete: true`
- **Use `false`** only after reviewing WhatIf output

**Recommended workflow**:
1. Run with `whatif: true` (preview)
2. Review console output
3. Set `whatif: false` to actually delete

```yaml
# Step 1: Preview
sync_delete: true
sync_delete_whatif: true   # Shows what would be deleted

# Step 2: Execute (after review)
sync_delete: true
sync_delete_whatif: false  # Actually deletes
```

#### `max_upload_workers` - Concurrent uploads

- **Default**: 4 (Graph API concurrent request limit)
- **Range**: 1-10 (capped to respect API limits)
- **When to adjust**:
  - Increase to 6-8 for high-bandwidth environments
  - Decrease to 2-3 if experiencing throttling
  - Keep at 4 for most use cases

```yaml
max_upload_workers: 4   # Default (recommended)
max_upload_workers: 8   # High-performance networks
max_upload_workers: 2   # Conservative (low throttling risk)
```

#### `max_hash_workers` - Hash calculation workers

- **Default**: Auto-detected (CPU count)
- **Range**: 1-unlimited
- **When to adjust**:
  - Usually no need to configure (auto-detection is optimal)
  - Decrease if running in resource-constrained environment
  - Leave empty for auto-detection

```yaml
max_hash_workers: ""    # Auto-detect (recommended)
max_hash_workers: 2     # Limit for low-resource containers
```

#### `login_endpoint` / `graph_endpoint` - Cloud environments

**Default**: Commercial cloud (`login.microsoftonline.com`, `graph.microsoft.com`)

**Government/Sovereign clouds**:

```yaml
# US Government GCC High
login_endpoint: "login.microsoftonline.us"
graph_endpoint: "graph.microsoft.us"

# Germany
login_endpoint: "login.microsoftonline.de"
graph_endpoint: "graph.microsoft.de"

# China (21Vianet)
login_endpoint: "login.chinacloudapi.cn"
graph_endpoint: "microsoftgraph.chinacloudapi.cn"
```

⚠️ **Endpoints must match the same cloud environment.**

</details>

<details>
<summary><strong>🎯 File Glob Patterns Reference</strong></summary>

### Pattern Syntax

| Pattern | Matches | Examples |
|---------|---------|----------|
| `*.pdf` | PDFs in root | `report.pdf`, `guide.pdf` |
| `**/*.pdf` | PDFs anywhere | `docs/report.pdf`, `archive/2024/guide.pdf` |
| `docs/*` | Files directly in docs | `docs/readme.md`, `docs/config.json` |
| `docs/**/*` | All files in docs tree | `docs/api/spec.yaml`, `docs/guides/intro.md` |
| `**/*.{md,txt}` | Markdown and text files | `readme.md`, `notes.txt` |
| `**/*` | Everything | All files and folders |

### Pattern Rules

- `*` - Matches any characters in single directory
- `**` - Matches zero or more directories
- `{a,b}` - Matches either a or b
- `!pattern` - Excludes matching files
- Always use forward slashes `/` (even on Windows)

### Common Patterns

```yaml
# All Markdown documentation
file_path: "**/*.md"

# Specific folder only
file_path: "reports/*"

# Multiple file types
file_path: "**/*.{pdf,docx,xlsx}"

# Everything except tests
file_path: "**/*"
exclude_patterns: "*.test.*,__tests__"
```

</details>

<details>
<summary><strong>🚫 Exclusion Patterns Reference</strong></summary>

### How Exclusions Work

Provide comma-separated patterns in `exclude_patterns`:

```yaml
exclude_patterns: "*.log,*.tmp,__pycache__,node_modules"
```

### Pattern Types

| Type | Example | Excludes |
|------|---------|----------|
| File extension | `*.log` | All `.log` files |
| Exact name | `config.json` | Files named `config.json` |
| Directory | `__pycache__` | Any `__pycache__` directory |
| Wildcard | `temp*` | Files starting with `temp` |
| Multiple | `*.pyc,*.pyo` | Both `.pyc` and `.pyo` |

### Matching Rules

- **Basename**: `*.tmp` matches `file.tmp`
- **Path component**: `node_modules` excludes `src/node_modules/file.js`
- **Case**: Sensitive on Linux, insensitive on Windows
- **Extensions**: `log` auto-expands to `*.log`

### Common Examples

**Python projects:**
```yaml
exclude_patterns: "*.pyc,*.pyo,__pycache__,.pytest_cache,.mypy_cache"
```

**Node.js projects:**
```yaml
exclude_patterns: "node_modules,*.log,dist,build,.cache"
```

**Build artifacts:**
```yaml
exclude_patterns: "*.dll,*.exe,*.so,bin,obj"
```

**Sensitive data:**
```yaml
exclude_patterns: ".env,.env.local,secrets.json,*.key,*.pem"
```

**Temp/hidden files:**
```yaml
exclude_patterns: "*.tmp,*.bak,.DS_Store,Thumbs.db,.git"
```

</details>

## 📚 Usage Examples

### Example 1: Sync Entire Repository

```yaml
- name: Checkout Repository
  uses: actions/checkout@v4

- name: Sync All Files
  uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    file_path: "**/*"
    file_path_recursive_match: true
    host_name: "company.sharepoint.com"
    site_name: "RepoDocumentation"
    upload_path: "Shared Documents/GitHub Mirror"
    tenant_id: ${{ secrets.SHAREPOINT_TENANT_ID }}
    client_id: ${{ secrets.SHAREPOINT_CLIENT_ID }}
    client_secret: ${{ secrets.SHAREPOINT_CLIENT_SECRET }}
```

### Example 2: Markdown Documentation Only

```yaml
- name: Checkout Repository
  uses: actions/checkout@v4

- name: Sync Documentation
  uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    file_path: "**/*.md"
    file_path_recursive_match: true
    convert_md_to_html: true  # Converts to styled HTML
    host_name: "company.sharepoint.com"
    site_name: "KnowledgeBase"
    upload_path: "Documentation"
    tenant_id: ${{ secrets.SHAREPOINT_TENANT_ID }}
    client_id: ${{ secrets.SHAREPOINT_CLIENT_ID }}
    client_secret: ${{ secrets.SHAREPOINT_CLIENT_SECRET }}
```

### Example 3: Exclude Build Artifacts

```yaml
- name: Checkout Repository
  uses: actions/checkout@v4

- name: Sync Clean Repository
  uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    file_path: "**/*"
    file_path_recursive_match: true
    exclude_patterns: "*.pyc,__pycache__,node_modules,*.log,.git,dist,build"
    host_name: "company.sharepoint.com"
    site_name: "DevDocs"
    upload_path: "Projects/MyApp"
    tenant_id: ${{ secrets.SHAREPOINT_TENANT_ID }}
    client_id: ${{ secrets.SHAREPOINT_CLIENT_ID }}
    client_secret: ${{ secrets.SHAREPOINT_CLIENT_SECRET }}
```

### Example 4: GovCloud Deployment

```yaml
- name: Checkout Repository
  uses: actions/checkout@v4

- name: Sync to GovCloud
  uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    file_path: "compliance/**/*"
    host_name: "agency.sharepoint.us"
    site_name: "Compliance"
    upload_path: "FY2025"
    tenant_id: ${{ secrets.SHAREPOINT_TENANT_ID }}
    client_id: ${{ secrets.SHAREPOINT_CLIENT_ID }}
    client_secret: ${{ secrets.SHAREPOINT_CLIENT_SECRET }}
    login_endpoint: "login.microsoftonline.us"
    graph_endpoint: "graph.microsoft.us"
```

### Example 5: Bidirectional Sync with Cleanup

```yaml
- name: Checkout Repository
  uses: actions/checkout@v4

- name: Full Sync with Deletion
  uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    file_path: "docs/**/*"
    file_path_recursive_match: true
    sync_delete: true              # Enable deletion
    sync_delete_whatif: true       # Preview first (safe)
    host_name: "company.sharepoint.com"
    site_name: "Documentation"
    upload_path: "Shared Documents/Docs"
    tenant_id: ${{ secrets.SHAREPOINT_TENANT_ID }}
    client_id: ${{ secrets.SHAREPOINT_CLIENT_ID }}
    client_secret: ${{ secrets.SHAREPOINT_CLIENT_SECRET }}
```

## 🔧 Advanced Features

### Parallel Processing for Maximum Performance

The action uses **parallel processing by default** for maximum performance with all file operations executing concurrently.

**Performance Improvements:**
- ⚡ **4-6x faster** overall sync time for typical repositories
- 🚀 **5-10x faster** uploads for repos with 50+ files
- 💨 **2-3x faster** hash calculation (utilizes all CPU cores)
- 📊 **10-20x fewer** API calls (batch metadata updates)

**What Runs in Parallel:**

| Operation | Workers | Description |
|-----------|---------|-------------|
| **File Uploads** | 4 (configurable) | Multiple files upload simultaneously |
| **Hash Calculation** | CPU count (auto) | Uses all available cores for ultra-fast hashing |
| **Markdown Conversion** | 4 (internal) | Multiple Mermaid diagrams render concurrently |
| **Metadata Updates** | Batched (20/request) | Graph API batch endpoint reduces API calls |

**Auto-Configuration:**

The action automatically detects your system specifications and configures optimal worker counts:
- Upload workers: Default 4 (respects Graph API concurrent limit)
- Hash workers: Detects CPU count for optimal parallel hashing
- Markdown workers: Fixed at 4 (balances performance vs memory)

**Thread Safety:**

All parallel operations are fully thread-safe with:
- Sequential console output (no garbled messages)
- Atomic statistics updates (accurate counts)
- Proper rate limiting coordination
- Batch queue for efficient metadata updates

**Console Output:**
```
============================================================
[✓] SYSTEM CONFIGURATION
============================================================
CPU Cores Available:       8
Upload Workers:            4 (concurrent uploads)
Hash Workers:              8 (parallel hashing)
Markdown Workers:          4 (parallel conversion)
Batch Metadata Updates:    Enabled (20 items/batch)
============================================================
```

**Configuration:**

```yaml
# Use defaults (recommended for most cases)
- uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    # ... other parameters

# Adjust for high-performance environments
- uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    max_upload_workers: 8      # Increase concurrent uploads
    max_hash_workers: 16       # Override auto-detection
    # ... other parameters

# Conservative settings (minimize throttling risk)
- uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    max_upload_workers: 2      # Reduce concurrent requests
    # ... other parameters
```

**Backward Compatibility:**

Parallel processing is fully compatible with existing workflows:
- ✅ Same console output format
- ✅ Same statistics structure
- ✅ Same error handling and retry logic
- ✅ No configuration changes required

**Performance Benchmarks:**

| Scenario | Sequential (Old) | Parallel (New) | Improvement |
|----------|-----------------|----------------|-------------|
| 100 files, 50 MB | 250s (4 min) | 40-60s (1 min) | **4-6x faster** |
| 500 files, 200 MB | 1200s (20 min) | 180-300s (3-5 min) | **4-6x faster** |

### Smart Sync with Content Hashing

The action uses **xxHash128** for lightning-fast change detection:

**How it works:**
1. Calculates hash for local file (3-6 GB/s processing speed)
2. Compares with `FileHash` column in SharePoint
3. Skips file if hash matches (unchanged)
4. Uploads only new or modified files

**Fallback:** Uses file size comparison if hash unavailable.

**Benefits:**
- ⚡ Typically skips 60-90% of files
- 💾 Saves bandwidth and time
- 🎯 More reliable than timestamps (especially in Docker/CI)
- 📊 Detailed statistics show comparison methods and hash operations

**Statistics Tracked:**
- **Comparison Methods**: Shows how many files were compared using hash vs size
- **FileHash Operations**: Tracks new hashes saved, hashes updated, and hash matches
- **Efficiency Metrics**: Displays bandwidth saved and skip rate

**Example Output:**
```
[✓] Smart sync enabled - unchanged files will be skipped
[✓] FileHash column available for hash-based comparison

Processing files...
[=] File unchanged (hash match): README.md
[*] File changed (hash mismatch): config.json
[+] New file to upload: changelog.md

============================================================
[✓] SYNC PROCESS COMPLETED
============================================================
[STATS] Sync Statistics:
   - New files uploaded:          255
   - Files updated:               129
   - Files skipped (unchanged):   569
   - Total files processed:       953

[COMPARE] File Comparison Methods:
   - Compared by hash:            682 (71.7%)
   - Compared by size:            271 (28.3%)

[HASH] FileHash Column Statistics:
   - New hashes saved:            255
   - Hashes updated:              129
   - Hash matches (skipped):      569

[DATA] Transfer Summary:
   - Data uploaded:   93.7 MB
   - Data skipped:    177.5 MB (569 files not re-uploaded)
   - Sync efficiency: 65.4% (bandwidth saved by smart sync)
============================================================
```

### Markdown Conversion

Converts `.md` files to GitHub-flavored HTML with embedded styling:

**Supported Features:**
- ✅ Headers (H1-H6) with anchors
- ✅ Tables with styling
- ✅ Code blocks with syntax highlighting
- ✅ Task lists with checkboxes
- ✅ Blockquotes
- ✅ Links and images
- ✅ **Mermaid diagrams with automatic mermaid sanitization** (rendered as embedded SVG)

<details>
<summary><strong>🔧 Automatic Mermaid Sanitization</strong></summary>

The action automatically fixes common Mermaid syntax issues before rendering.

### Before Sanitization (Would Fail)
````markdown
```mermaid
graph TD
    A[Server #1] --> B{Version |No| Skip}
    B -->||"Proceed"|| C[Deploy]
    C --> end
    D[Post Action<br/>(Restart)]
```
````

### After Sanitization (Will Render)

````markdown
```mermaid
graph TD
    A[Server &#35;1] --> B{Version &#124;No&#124; Skip}
    B -->|'Proceed'| C[Deploy]
    C --> End
    D[Post Action<br>(Restart)]
```
````

### What Gets Fixed

| Issue | Before | After | Why |
|-------|--------|-------|-----|
| Special chars in nodes | `[Server #1]` | `[Server &#35;1]` | `#` breaks parser |
| Pipes in diamonds | `{Version \|No\| Skip}` | `{Version &#124;No&#124; Skip}` | Unescaped pipes break syntax |
| Comparison operators | `{Version >= 1.0?}` | `{Version &#62;= 1.0?}` | `<` and `>` break diamond nodes |
| Double pipes in edges | `-->\|\|"Proceed"\|\|` | `-->\|'Proceed'\|` | Invalid edge syntax |
| Double quotes | `\|"text"\|` | `\|'text'\|` | Prevents syntax errors |
| Reserved words | `end` | `End` | Lowercase breaks flowcharts |
| XHTML tags | `<br/>` | `<br>` | Mermaid only supports `<br>` |

**Supported Node Shapes:**
- Square brackets `[text]`
- Parentheses `(text)`, `((text))`
- Diamond/rhombus `{text}`, `{{text}}`
- Trapezoids `[/text\]`, `[\text/]`

All shapes are auto-sanitized for special characters. No manual fixes needed!

</details>

### Sync Deletion

**Bidirectional sync** - automatically removes SharePoint files that no longer exist in your repository.

**Use Cases:**
- 🔄 Renamed files (old name removed)
- 🗑️ Deleted files (removed from SharePoint)
- 📁 Restructured folders (cleanup old locations)
- 📝 Obsolete documentation (auto-removal)

**Safety Features:**
1. **Explicit opt-in** (`sync_delete: false` by default)
2. **WhatIf mode** (preview deletions before executing)
3. **Scoped deletion** (only within `upload_path`)
4. **Statistics** (audit trail of deletions)

**Configuration:**

| Parameter | Default | Purpose |
|-----------|---------|---------|
| `sync_delete` | `false` | Enable deletion feature |
| `sync_delete_whatif` | `true` | Preview mode (safe default) |

**Recommended Workflow:**

**Step 1: Preview (Safe)**
```yaml
sync_delete: true
sync_delete_whatif: true  # Shows what would be deleted
```

**Console Output:**
```
[!] Sync deletion enabled in WHATIF mode
[*] Found 3 orphaned files (no actual deletions)

File Deleted (WhatIf): old-readme.md
File Deleted (WhatIf): deprecated/guide.md
File Deleted (WhatIf): archive/notes.txt

[✓] WhatIf: Would delete 3 files
```

**Step 2: Execute (After Review)**
```yaml
sync_delete: true
sync_delete_whatif: false  # Actually deletes files
```

**Console Output:**
```
[!] Sync deletion enabled - files will be DELETED
[*] Found 3 orphaned files to delete

File Deleted: old-readme.md
File Deleted: deprecated/guide.md
File Deleted: archive/notes.txt

[✓] Successfully deleted 3 files
```

⚠️ **Important:** Always test with WhatIf first. Deleted files may be in SharePoint recycle bin (depends on configuration).

💡 **Best Practice:** Use in scheduled workflows for automated cleanup.

#### 🔍 Troubleshooting Sync Deletion

If sync deletion is marking unexpected files for deletion, enable **DEBUG mode** to see detailed comparison:

```yaml
- name: Sync with Debug Output
  uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    # ... your parameters
    sync_delete: true
    sync_delete_whatif: true
  env:
    DEBUG: "true"  # Enable detailed debug logging
```

**Debug output shows:**
- Full list of SharePoint files found
- Full list of local files in sync set
- Path-by-path comparison results
- Whether items are files or folders
- Exact reason each file is orphaned or matched

This helps diagnose path mismatches, folder structure issues, or markdown conversion problems.

### Filename Sanitization

SharePoint restricts certain characters. The action auto-converts them:

| Invalid | Replacement |
|---------|-------------|
| `#`, `%`, `&` | Fullwidth Unicode equivalents |
| `:`, `<`, `>`, `?`, `\|` | Fullwidth Unicode equivalents |

Reserved names (CON, PRN, AUX, NUL, etc.) are prefixed with underscore.

**Example:** `file:name#test.md` → `file：name＃test.md`

## 🔒 Security

### Best Practices

- ✅ **Never commit secrets** to repository
- ✅ **Use GitHub Secrets** for all credentials
- ✅ **Use Sites.Selected** for granular access control (recommended)
- ✅ **Limit permissions** to only required sites
- ✅ **Review app permissions** regularly
- ✅ **Use branch protection** to control workflow triggers
- ✅ **Rotate client secrets** before expiration

### Sites.Selected Setup

**Enhanced security** - grant app access to specific sites only (vs. tenant-wide access).

<details>
<summary><strong>🔐 Sites.Selected Configuration Guide</strong></summary>

#### 1. Configure App Registration

In Azure AD, grant **Sites.Selected** permission:
- API: Microsoft Graph
- Permission: Sites.Selected (Application)
- Type: Application permission

#### 2. Grant Site Access

**Method A: Graph API**
```bash
# Get site ID
GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{sitename}

# Grant app access (requires admin permission)
POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
{
  "roles": ["write"],
  "grantedToIdentities": [{
    "application": {
      "id": "{client-id}",
      "displayName": "{app-name}"
    }
  }]
}
```

**Method B: PowerShell (PnP)**
```powershell
# Install PnP PowerShell
Install-Module -Name PnP.PowerShell

# Connect to site
Connect-PnPOnline -Url "https://{tenant}.sharepoint.com/sites/{sitename}" -Interactive

# Grant permissions
Grant-PnPAzureADAppSitePermission `
  -AppId "{client-id}" `
  -DisplayName "{app-name}" `
  -Permissions Write

# Verify
Get-PnPAzureADAppSitePermission
```

#### 3. Use in Workflow

No workflow changes needed! Works identically:

```yaml
- uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    site_name: "TeamSite"  # Must match granted site
    # ... other parameters (same as before)
```

**Benefits:**
- ✅ App can only access specified sites
- ✅ Compliance-friendly granular control
- ✅ Clear audit trail
- ✅ No workflow changes required

</details>

## 🐛 Troubleshooting

<details>
<summary><strong>Common Issues & Solutions</strong></summary>

### 1. Authentication Failed

**Error:** `(401) Unauthorized`

**Solutions:**
- ✅ Verify Client ID and Secret are correct
- ✅ Check app has required permissions:
  - `Sites.ReadWrite.All` OR `Sites.Selected`
  - `Sites.Manage.All` (optional, for FileHash column)
- ✅ For Sites.Selected: Verify app granted access to site
- ✅ Ensure Tenant ID matches SharePoint instance

### 2. File Not Found

**Error:** `No files matched pattern`

**Solutions:**
- ✅ Check glob pattern syntax
- ✅ Enable `file_path_recursive_match: true` for nested directories
- ✅ Add debug step: `run: ls -la` to verify files exist
- ✅ Check `exclude_patterns` isn't excluding your files

### 3. Upload Timeout

**Error:** `Operation timed out`

**Solutions:**
- ✅ Large files automatically retry (check logs)
- ✅ Consider splitting large uploads into multiple actions
- ✅ Increase `max_retries` for unstable networks
- ✅ Check network connectivity

### 4. Markdown Conversion Failed

**Error:** `Mermaid conversion failed`

**Solutions:**
- ✅ Verify Mermaid syntax (use [Mermaid Live Editor](https://mermaid.live))
- ✅ Simplify complex diagrams
- ✅ Check for unsupported Mermaid features
- ✅ Automatic sanitization handles most issues

### 5. Permission Denied (Column Creation)

**Error:** `Access denied creating FileHash column`

**Solutions:**
- ✅ This is expected if app lacks `Sites.Manage.All`
- ✅ Action automatically falls back to file size comparison
- ✅ Grant `Sites.Manage.All` for hash-based comparison (optional)

### Debug Steps

Add before upload step to inspect files:

```yaml
- name: Debug - List Files
  run: |
    echo "Files matching pattern:"
    ls -la docs/**/*.md
```

Enable verbose logging:
```yaml
- name: Upload with Debug
  uses: AunalyticsManagedServices/sharepoint-file-upload-action@v3
  with:
    # ... parameters
  env:
    ACTIONS_STEP_DEBUG: true
```

</details>

## 📊 Performance

### Optimization Tips

1. **Use specific glob patterns** - `"docs/*.md"` vs. `"**/*"`
2. **Enable smart sync** (default) - skips unchanged files
3. **Upload frequently** - minimizes changes per sync
4. **Organize files logically** - enables targeted patterns
5. **Schedule off-peak** - for large syncs

**Hashing Performance:** xxHash128 processes files at 3-6 GB/s (10-20x faster than SHA-256).

<details>
<summary><strong>Performance Benchmarks</strong></summary>

| Files | Size | Smart Sync | Force Upload |
|-------|------|------------|--------------|
| 10 | 50 MB | ~15 seconds | ~45 seconds |
| 100 | 500 MB | ~30 seconds | ~5 minutes |
| 1000 | 5 GB | ~2 minutes | ~30 minutes |

*Times vary based on network speed and file changes. Smart sync typically skips 60-90% of files.*

**File Size Handling:**
- Small files (<4MB): Direct upload (~10 files/second)
- Large files (≥4MB): Chunked upload (~50 MB/second)

</details>

## 🤝 Contributing

We welcome contributions! Here's how to get started:

### Development Setup

1. Fork the repository
2. Create feature branch: `git checkout -b feature/amazing-feature`
3. Make your changes
4. Test locally (see below)
5. Commit: `git commit -m 'Add amazing feature'`
6. Push: `git push origin feature/amazing-feature`
7. Open Pull Request

### Local Testing

```bash
# Build Docker image
docker build -t sharepoint-upload .

# Test with sample files
docker run sharepoint-upload \
  "test-site" \
  "test.sharepoint.com" \
  "tenant-id" \
  "client-id" \
  "client-secret" \
  "test-path" \
  "*.md"
```

## 📞 Support

- **🐛 Issues:** [GitHub Issues](https://github.com/AunalyticsManagedServices/sharepoint-file-upload-action/issues)
- **💬 Discussions:** [GitHub Discussions](https://github.com/AunalyticsManagedServices/sharepoint-file-upload-action/discussions)
- **📧 Email:** [support@aunalytics.com](mailto:support@aunalytics.com)

## 📜 License

This project is licensed under the MIT License - see [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

Built with these excellent open-source projects:
- [Office365-REST-Python-Client](https://github.com/vgrem/Office365-REST-Python-Client) - SharePoint/Graph API
- [Mistune](https://github.com/lepture/mistune) - Markdown parsing
- [Mermaid-CLI](https://github.com/mermaid-js/mermaid-cli) - Diagram rendering
- [xxHash](https://github.com/Cyan4973/xxHash) - Ultra-fast hashing

---

**Made with ❤️ by [Aunalytics Managed Services](https://www.aunalytics.com)**
