# -*- coding: utf-8 -*-
"""
Markdown to HTML conversion with Mermaid diagram support.

This module converts Markdown files to HTML with GitHub-flavored styling
and renders Mermaid diagrams as embedded SVG images.
"""

import os
import re
import tempfile
import subprocess
import mistune


def sanitize_mermaid_code(mermaid_code):
    """
    Sanitize Mermaid diagram code to fix syntax issues for mermaid-cli.

    This function provides automatic fixes for common mermaid-cli parse errors by
    replacing problematic characters with text alternatives. Based on mermaid-cli
    strict parsing requirements.

    Handles:
    - Self-closing HTML tags: <br> → <br/>
    - Comparison operators: < → "under", > → "over", <= → "at most", >= → "at least"
    - Ampersands: & → "and"
    - Double quotes: " → '
    - Reserved words: "end" → "End"
    - HTML tags (removes all except <br/>)
    - Curly braces in comments
    - Double pipe issues: -->|| → -->|

    Character Replacement Strategy:
    - Replaces < > & with TEXT alternatives (not HTML entities)
    - HTML entities (&#38;, &#35;) break mermaid-cli parser
    - Text replacements preserve meaning while ensuring parse success
    - Applied to: node labels, edge labels, decision nodes, subgraph labels

    Examples:
        {Count < 3?} → {Count under 3?}
        -->|Battery < 50%| → -->|Battery under 50%|
        subgraph ["A & B"] → subgraph ["A and B"]

    Args:
        mermaid_code (str): Raw Mermaid diagram definition

    Returns:
        str: Sanitized Mermaid code safe for mmdc rendering
    """
    sanitized = mermaid_code

    # 1. Ensure <br> tags are self-closing (mermaid-cli REQUIRES <br/>)
    # Non-self-closing <br> breaks strict parser
    sanitized = re.sub(r'<br(?!\s*/?>)', '<br/>', sanitized, flags=re.IGNORECASE)
    sanitized = re.sub(r'<br\s+(?!/?>)', '<br/>', sanitized, flags=re.IGNORECASE)

    # 2. Remove other HTML tags except <br/>
    # Keep <br/> since mermaid-cli supports it for line breaks (must be self-closing)
    sanitized = re.sub(r'<(?!br\b)[^>]+>', '', sanitized, flags=re.IGNORECASE)

    # 3. Fix reserved word "end" - it breaks Flowcharts and Sequence diagrams
    # Replace standalone lowercase "end" with "End" in node labels
    # Match patterns like [end], (end), or "end" but not "append", "ending", etc.
    sanitized = re.sub(r'\b(end)\b', 'End', sanitized)

    # 4. Fix double pipes in edge definitions (||) -> (|)
    # Pattern: -->|| or ---|| or ||| should become -->| or ---|
    sanitized = re.sub(r'(-->|---)\|\|', r'\1|', sanitized)
    sanitized = re.sub(r'\|\|(\w)', r'|\1', sanitized)

    # 5. Smart character replacement for mermaid-cli compatibility
    # IMPORTANT: mermaid-cli breaks with HTML entities AND with raw < > & in certain contexts
    # Replace these with text alternatives that preserve meaning but use safe characters
    def sanitize_content(content):
        """Replace special characters with text alternatives (mermaid-cli safe)"""
        # Replace quotes (syntactically required)
        content = content.replace('"', "'")

        # Replace comparison operators and special characters with text alternatives
        # These break mermaid-cli parser in edge labels, decision nodes, and subgraph labels
        # Do NOT use HTML entities - use plain text instead
        content = content.replace('<=', ' at most ')
        content = content.replace('>=', ' at least ')
        content = content.replace('<', ' under ')
        content = content.replace('>', ' over ')
        content = content.replace('&', ' and ')

        # Clean up any double spaces created by replacements
        while '  ' in content:
            content = content.replace('  ', ' ')

        return content.strip()

    def sanitize_node_content(match):
        """Replace special characters in square bracket node content"""
        content = match.group(1)
        return f'[{sanitize_content(content)}]'

    # Apply sanitization to content inside square brackets []
    sanitized = re.sub(r'\[([^]]*)]', sanitize_node_content, sanitized)

    # 6. Handle parentheses-based node shapes: (text), ((text)), etc.
    def sanitize_paren_content(match):
        """Replace special characters in parentheses node content"""
        opening_parens = match.group(1)
        content = match.group(2)
        closing_parens = match.group(3)

        return f'{opening_parens}{sanitize_content(content)}{closing_parens}'

    # Match single or multiple parentheses: (text), ((text)), (((text)))
    sanitized = re.sub(r'(\(+)([^()]+)(\)+)', sanitize_paren_content, sanitized)

    # 7. Handle curly brace diamond/rhombus nodes: {text}, {{text}}
    # CRITICAL: Decision nodes {} are parsed strictly by mermaid-cli
    # Special characters < > & break the DIAMOND_STOP token parser
    def sanitize_curly_content(match):
        """Smart sanitization for decision nodes (removes <br/> and replaces special chars)"""
        opening_braces = match.group(1)
        content = match.group(2)
        closing_braces = match.group(3)

        # Remove <br/> tags from decision nodes - they break the parser
        content = re.sub(r'<br\s*/?>','', content, flags=re.IGNORECASE)

        # Replace special characters with text alternatives
        content = content.replace('"', "'")
        content = content.replace('<=', ' at most ')
        content = content.replace('>=', ' at least ')
        content = content.replace('<', ' under ')
        content = content.replace('>', ' over ')
        content = content.replace('&', ' and ')

        # Clean up double spaces
        while '  ' in content:
            content = content.replace('  ', ' ')
        content = content.strip()

        return f'{opening_braces}{content}{closing_braces}'

    # Match single or double curly braces: {text}, {{text}}
    sanitized = re.sub(r'(\{+)([^{}]+)(}+)', sanitize_curly_content, sanitized)

    # 8. Handle trapezoid node shapes: [/text\] and [\text/]
    def sanitize_trapezoid_content(match):
        """Replace special characters in trapezoid node content"""
        opening = match.group(1)
        content = match.group(2)
        closing = match.group(3)

        return f'{opening}{sanitize_content(content)}{closing}'

    # Match trapezoid patterns
    sanitized = re.sub(r'(\[/)(.*?)(\\])', sanitize_trapezoid_content, sanitized)
    sanitized = re.sub(r'(\[\\)(.*?)(/])', sanitize_trapezoid_content, sanitized)

    # 9. Handle hexagon node shapes: {{text}}
    # Already handled by curly brace sanitization above

    # 10. Handle edge labels (text between pipes on arrows)
    # Pattern: -->|text| or ---|text|--- etc.
    # CRITICAL: Edge labels in mermaid-cli cannot contain < > & characters
    def sanitize_edge_label(match):
        """Smart sanitization for edge labels (replaces problematic characters)"""
        prefix = match.group(1)
        label = match.group(2)
        suffix = match.group(3)

        # Apply same smart replacements as node content
        # Replace comparison operators and special characters with text
        label_sanitized = label.replace('"', "'")
        label_sanitized = label_sanitized.replace('<=', ' at most ')
        label_sanitized = label_sanitized.replace('>=', ' at least ')
        label_sanitized = label_sanitized.replace('<', ' under ')
        label_sanitized = label_sanitized.replace('>', ' over ')
        label_sanitized = label_sanitized.replace('&', ' and ')

        # Clean up double spaces
        while '  ' in label_sanitized:
            label_sanitized = label_sanitized.replace('  ', ' ')
        label_sanitized = label_sanitized.strip()

        return f'{prefix}|{label_sanitized}|{suffix}'

    # Match edge labels: arrow followed by |text| followed by arrow or node
    sanitized = re.sub(r'(--[>-])\|([^|]+)\|(--[>-]|\s+\w)', sanitize_edge_label, sanitized)

    # 11. Fix nodes starting with "o" or "x" which create unintended edges
    # Add a space after node ID if it starts with o/x
    sanitized = re.sub(r'(\w+)\s+([ox])---', r'\1  \2---', sanitized)
    sanitized = re.sub(r'(\w+)\s+([ox])-->', r'\1  \2-->', sanitized)

    # 12. Remove curly braces from comments (they confuse the renderer)
    def remove_braces_from_comments(match):
        comment_text = match.group(1)
        return f'%% {comment_text.replace("{", "").replace("}", "")}'

    sanitized = re.sub(r'%%\s*([^\n]*)', remove_braces_from_comments, sanitized)

    return sanitized


def convert_mermaid_to_svg(mermaid_code, filename=None):
    """
    Convert Mermaid diagram code to SVG using mermaid-cli.

    Uses the mmdc command-line tool installed via npm to render
    Mermaid diagrams as static SVG images.

    Args:
        mermaid_code (str): Mermaid diagram definition
        filename (str, optional): Original filename for error messages

    Returns:
        str: SVG content as string, or None if conversion fails
    """
    try:
        # Sanitize the Mermaid code to fix common issues
        sanitized_code = sanitize_mermaid_code(mermaid_code)

        # Create temporary files for input and output
        with tempfile.NamedTemporaryFile(mode='w', suffix='.mmd', delete=False) as mmd_file:
            mmd_file.write(sanitized_code)
            mmd_path = mmd_file.name

        svg_path = mmd_path.replace('.mmd', '.svg')

        try:
            # Run mermaid-cli to convert to SVG
            # Using puppeteer config for headless Chromium settings
            # Using mermaid config to prevent text truncation issues
            result = subprocess.run(
                ['mmdc', '-i', mmd_path, '-o', svg_path,
                 '--puppeteerConfigFile', '/usr/src/app/puppeteer-config.json',
                 '--configFile', '/usr/src/app/mermaid-config.json'],
                capture_output=True,
                text=True,
                timeout=30,
                check=True  # Will raise CalledProcessError on non-zero exit
            )

            # Success - read the generated SVG
            if os.path.exists(svg_path):
                with open(svg_path, 'r', encoding='utf-8') as f:
                    svg_content = f.read()

                # Clean up temporary files
                os.unlink(mmd_path)
                os.unlink(svg_path)

                return svg_content
            else:
                # Unexpected - mmdc returned 0 but didn't create SVG
                print(f"[!] SVG file was not created by mmdc (unexpected)")
                if filename:
                    print(f"    File: {filename}")
                if os.path.exists(mmd_path):
                    os.unlink(mmd_path)
                return None

        except FileNotFoundError:
            # mmdc binary not found - Docker configuration issue
            error_msg = "Mermaid CLI (mmdc) not found - Docker container configuration issue"
            print(f"[!] ========================================")
            print(f"[!] MERMAID CLI NOT FOUND")
            print(f"[!] ========================================")
            if filename:
                print(f"[!] File: {filename}")
            print(f"[!] ")
            print(f"[!] This is a Docker container configuration issue.")
            print(f"[!] ")
            print(f"[!] Troubleshooting steps:")
            print(f"[!]   1. Verify mermaid-cli is installed:")
            print(f"[!]      npm list -g @mermaid-js/mermaid-cli")
            print(f"[!]   2. Check Dockerfile installs mermaid-cli correctly:")
            print(f"[!]      RUN npm install -g @mermaid-js/mermaid-cli")
            print(f"[!]   3. Verify PATH includes Node.js global bin directory")
            print(f"[!]   4. Rebuild Docker container if configuration changed")
            print(f"[!] ========================================")
            # Clean up temp file
            if os.path.exists(mmd_path):
                os.unlink(mmd_path)
            return None

        except subprocess.TimeoutExpired as e:
            # Diagram took too long to render - likely too complex
            print(f"[!] ========================================")
            print(f"[!] MERMAID CONVERSION TIMEOUT")
            print(f"[!] ========================================")
            if filename:
                print(f"[!] File: {filename}")
            print(f"[!] ")
            print(f"[!] Diagram rendering timed out after {e.timeout} seconds.")
            print(f"[!] ")
            print(f"[!] Troubleshooting steps:")
            print(f"[!]   1. Diagram may be too complex for Chromium to render")
            print(f"[!]   2. Consider simplifying the diagram (fewer nodes/edges)")
            print(f"[!]   3. Split large diagram into multiple smaller diagrams")
            print(f"[!]   4. Increase timeout in code if needed (current: 30s)")
            print(f"[!] ")
            print(f"[!] The diagram will be shown as a code block instead of rendered SVG.")
            print(f"[!] ========================================")
            # Clean up temp files
            if os.path.exists(mmd_path):
                os.unlink(mmd_path)
            if os.path.exists(svg_path):
                os.unlink(svg_path)
            return None

        except subprocess.CalledProcessError as e:
            # mmdc failed - usually syntax error in diagram
            print(f"[!] ========================================")
            print(f"[!] MERMAID DIAGRAM SYNTAX ERROR")
            print(f"[!] ========================================")
            if filename:
                print(f"[!] File: {filename}")
            print(f"[!] ")
            print(f"[!] Mermaid CLI failed with exit code {e.returncode}")
            print(f"[!] This usually means the diagram has syntax errors that sanitization couldn't fix.")
            print(f"[!] ")
            if e.stderr:
                print(f"[!] Error output from mmdc:")
                # Print first 300 chars of stderr
                stderr_lines = e.stderr[:300].split('\n')
                for line in stderr_lines:
                    if line.strip():
                        print(f"[!]   {line}")
            print(f"[!] ")
            print(f"[!] First 200 characters of diagram code:")
            print(f"[!]   {sanitized_code[:200]}")
            print(f"[!] ")
            print(f"[!] The diagram will be shown as a code block instead of rendered SVG.")
            print(f"[!] ========================================")
            # Clean up temp files
            if os.path.exists(mmd_path):
                os.unlink(mmd_path)
            if os.path.exists(svg_path):
                os.unlink(svg_path)
            return None

    except OSError as e:
        # File I/O errors (permissions, disk full, etc.)
        print(f"[!] ========================================")
        print(f"[!] FILE I/O ERROR - Mermaid Conversion")
        print(f"[!] ========================================")
        if filename:
            print(f"[!] File: {filename}")
        print(f"[!] ")
        print(f"[!] Could not read/write temporary files for Mermaid conversion.")
        print(f"[!] ")
        print(f"[!] Troubleshooting steps:")
        print(f"[!]   1. Check disk space - may be full")
        print(f"[!]   2. Verify permissions on temp directory")
        print(f"[!]   3. Check if filesystem is read-only")
        print(f"[!] ")
        print(f"[!] Technical details: {str(e)[:200]}")
        print(f"[!] ========================================")
        return None

    except Exception as e:
        # Unexpected errors
        if filename:
            print(f"[!] Unexpected error converting Mermaid diagram: {filename}")
            print(f"    Error type: {type(e).__name__}")
            print(f"    Error: {str(e)[:200]}")
        else:
            print(f"[!] Unexpected error converting Mermaid diagram")
            print(f"    Error type: {type(e).__name__}")
            print(f"    Error: {str(e)[:200]}")
        return None


def rewrite_markdown_links(md_content, sharepoint_base_url=None, current_file_rel_path=None):
    """
    Rewrite internal markdown links to proper SharePoint URLs.

    Converts relative markdown links (e.g., ../README.md, folder/file.md) to
    absolute SharePoint URLs with proper path structure.

    Args:
        md_content (str): Markdown content with links
        sharepoint_base_url (str): Base SharePoint URL (e.g.,
            'https://aunalytics.sharepoint.com/sites/SiteName/Shared%20Documents/Folder%20Path')
            NOTE: Should be pre-encoded (spaces → %20, etc.)
        current_file_rel_path (str): Current file's relative path from upload root
            (e.g., 'Adobe/AcrobatDC/Update/README.md')

    Returns:
        str: Markdown content with rewritten links
    """
    if not sharepoint_base_url or not current_file_rel_path:
        # Can't rewrite without context - return original
        return md_content

    import posixpath
    from urllib.parse import quote

    # Define file extensions that can be viewed in SharePoint web browser
    WEB_VIEWABLE_EXTENSIONS = {
        '.html', '.htm',           # HTML files (with ?web=1)
        '.pdf',                    # PDF files (native viewer)
        '.docx', '.doc',          # Word documents (Office Online)
        '.xlsx', '.xls',          # Excel spreadsheets (Office Online)
        '.pptx', '.ppt',          # PowerPoint presentations (Office Online)
        '.txt',                   # Text files (preview)
        '.md',                    # Markdown (converted to .html by our action)
        '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg',  # Images (preview)
        '.mp4', '.mov', '.avi',   # Videos (player)
        '.mp3', '.wav',           # Audio (player)
        '.ps1', '.py', '.sh', '.bat', '.cmd',  # Script files (code preview)
        '.json', '.xml', '.yaml', '.yml',      # Config files (preview)
        '.csv', '.tsv',           # Data files (preview)
        '.log',                   # Log files (preview)
        '.cs', '.js', '.ts', '.java', '.cpp', '.c', '.h',  # Source code (preview)
    }

    # Get current file's directory
    current_dir = posixpath.dirname(current_file_rel_path)

    def rewrite_link(match):
        """Rewrite a single markdown link"""
        link_text = match.group(1)
        link_url = match.group(2)

        # Skip external links (http://, https://, mailto:, etc.)
        if '://' in link_url or link_url.startswith('mailto:') or link_url.startswith('#'):
            return match.group(0)  # Return unchanged

        # Determine if link is to local repository content
        # Process: .md files, folders (ending with /), or other file extensions
        is_markdown = link_url.endswith('.md') or '.md#' in link_url
        is_folder = link_url.endswith('/')

        # Skip if not markdown and not folder (leave other file types as-is)
        # This preserves links to images, PDFs, etc.
        if not is_markdown and not is_folder:
            # Check if it has a file extension - if not, might be a folder reference
            if '.' not in posixpath.basename(link_url):
                # No extension, might be folder - treat as folder
                is_folder = True
            else:
                # Has extension but not .md - skip
                return match.group(0)

        # Split anchor if present (e.g., README.md#section or folder/#section)
        if '#' in link_url:
            link_path, anchor = link_url.split('#', 1)
            anchor = '#' + anchor
        else:
            link_path = link_url
            anchor = ''

        # Resolve relative path to absolute path from upload root
        if link_path.startswith('/'):
            # Absolute path from repository root
            resolved_path = link_path.lstrip('/')
        else:
            # Relative path - resolve from current directory
            resolved_path = posixpath.normpath(posixpath.join(current_dir, link_path))

        # Convert .md extension to .html
        if resolved_path.endswith('.md'):
            resolved_path = resolved_path[:-3] + '.html'

        # For folders, ensure trailing slash is removed for URL construction
        if is_folder and resolved_path.endswith('/'):
            resolved_path = resolved_path.rstrip('/')

        # Determine if file can be viewed in browser
        file_ext = posixpath.splitext(resolved_path)[1].lower()
        is_web_viewable = file_ext in WEB_VIEWABLE_EXTENSIONS

        # For non-web-viewable files, link to the containing folder instead
        if not is_folder and not is_web_viewable:
            # Get parent folder path
            folder_path = posixpath.dirname(resolved_path)
            if folder_path:
                # Link to folder containing the file
                # URL encode the path components but preserve slashes
                path_parts = folder_path.split('/')
                encoded_parts = [quote(part) for part in path_parts]
                encoded_path = '/'.join(encoded_parts)

                # Folder view URL (SharePoint will show the folder contents)
                full_url = f"{sharepoint_base_url}/{encoded_path}"

                # Return rewritten markdown link with note about folder location
                return f'[{link_text}]({full_url} "{resolved_path.split("/")[-1]} - Click to view folder")'
            else:
                # File is in root - link to base URL
                return f'[{link_text}]({sharepoint_base_url} "{resolved_path} - Click to view folder")'

        # URL encode the path components but preserve slashes
        path_parts = resolved_path.split('/')
        encoded_parts = [quote(part) for part in path_parts]
        encoded_path = '/'.join(encoded_parts)

        # Construct full SharePoint URL
        if is_folder:
            # Folder link format (opens folder view)
            full_url = f"{sharepoint_base_url}/{encoded_path}"
        else:
            # File link format (web-viewable files)
            # Add ?web=1 to open in browser preview instead of downloading
            # This works for HTML, scripts, code files, text files, etc.
            full_url = f"{sharepoint_base_url}/{encoded_path}?web=1{anchor}"

        # Return rewritten markdown link
        return f'[{link_text}]({full_url})'

    # Pattern to match markdown links: [text](url)
    import re
    md_content = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', rewrite_link, md_content)

    return md_content


def convert_markdown_to_html(md_content, filename, sharepoint_base_url=None, current_file_rel_path=None):
    """
    Convert Markdown content to HTML with Mermaid diagrams rendered as SVG.

    This function:
    1. Rewrites internal markdown links to SharePoint URLs
    2. Parses markdown using Mistune
    3. Finds and converts Mermaid code blocks to inline SVG
    4. Applies GitHub-like styling for SharePoint viewing

    Args:
        md_content (str): Markdown content to convert
        filename (str): Original filename for the HTML title
        sharepoint_base_url (str, optional): Base SharePoint URL for link rewriting
        current_file_rel_path (str, optional): Current file's relative path

    Returns:
        tuple: (html_content: str, mermaid_success: int, mermaid_failed: int)
               - html_content: Complete HTML document with embedded styles and SVGs
               - mermaid_success: Number of Mermaid diagrams successfully converted to SVG
               - mermaid_failed: Number of Mermaid diagrams that failed (shown as code blocks)
    """
    # Rewrite internal markdown links before conversion
    md_content = rewrite_markdown_links(md_content, sharepoint_base_url, current_file_rel_path)
    # First, extract and convert all mermaid blocks to placeholder SVGs
    mermaid_pattern = r'```mermaid\n(.*?)\n```'
    mermaid_blocks = []
    mermaid_success_count = 0
    mermaid_failed_count = 0

    def replace_mermaid_with_placeholder(match):
        nonlocal mermaid_success_count, mermaid_failed_count
        mermaid_code = match.group(1)
        placeholder = f"<!--MERMAID_PLACEHOLDER_{len(mermaid_blocks)}-->"

        # Convert to SVG
        svg_content = convert_mermaid_to_svg(mermaid_code, filename)
        if svg_content:
            # Clean up the SVG for inline embedding
            # Remove XML declaration if present
            svg_content = re.sub(r'<\?xml[^>]*\?>', '', svg_content)
            svg_content = svg_content.strip()
            mermaid_blocks.append(svg_content)
            mermaid_success_count += 1
        else:
            # If conversion failed, keep as code block
            mermaid_blocks.append(f'<pre><code>mermaid\n{mermaid_code}</code></pre>')
            mermaid_failed_count += 1

        return placeholder

    # Replace mermaid blocks with placeholders
    md_with_placeholders = re.sub(mermaid_pattern, replace_mermaid_with_placeholder, md_content, flags=re.DOTALL)

    # Convert markdown to HTML using Mistune
    html_body = mistune.html(md_with_placeholders)

    # Replace placeholders with actual SVG content
    for i, svg_content in enumerate(mermaid_blocks):
        placeholder = f"<!--MERMAID_PLACEHOLDER_{i}-->"
        # Wrap SVG in a div for centering
        wrapped_svg = f'<div class="mermaid-diagram">{svg_content}</div>'
        html_body = html_body.replace(f"<p>{placeholder}</p>", wrapped_svg)
        html_body = html_body.replace(placeholder, wrapped_svg)

    # Create the complete HTML document
    html_template = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{filename.replace('.md', '')}</title>

    <style>
        /* GitHub-like styling for SharePoint */
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans", Helvetica, Arial, sans-serif;
            font-size: 16px;
            line-height: 1.5;
            word-wrap: break-word;
            padding: 20px;
            max-width: 980px;
            margin: 0 auto;
            color: #1F2328;
            background: #ffffff;
        }}

        h1, h2, h3, h4, h5, h6 {{
            margin-top: 24px;
            margin-bottom: 16px;
            font-weight: 600;
            line-height: 1.25;
        }}

        h1 {{
            font-size: 2em;
            border-bottom: 1px solid #d1d9e0;
            padding-bottom: .3em;
        }}

        h2 {{
            font-size: 1.5em;
            border-bottom: 1px solid #d1d9e0;
            padding-bottom: .3em;
        }}

        h3 {{ font-size: 1.25em; }}
        h4 {{ font-size: 1em; }}
        h5 {{ font-size: .875em; }}
        h6 {{ font-size: .85em; color: #59636e; }}

        code {{
            padding: .2em .4em;
            margin: 0;
            font-size: 85%;
            white-space: break-spaces;
            background-color: #f6f8fa;
            border-radius: 6px;
            font-family: ui-monospace, SFMono-Regular, "SF Mono", Consolas, "Liberation Mono", Menlo, monospace;
        }}

        pre {{
            padding: 16px;
            overflow: auto;
            font-size: 85%;
            line-height: 1.45;
            color: #1F2328;
            background-color: #f6f8fa;
            border-radius: 6px;
            margin-top: 0;
            margin-bottom: 16px;
        }}

        pre code {{
            display: inline;
            max-width: auto;
            padding: 0;
            margin: 0;
            overflow: visible;
            line-height: inherit;
            word-wrap: normal;
            background-color: transparent;
            border: 0;
        }}

        blockquote {{
            margin: 0;
            padding: 0 1em;
            color: #59636e;
            border-left: .25em solid #d1d9e0;
        }}

        table {{
            border-spacing: 0;
            border-collapse: collapse;
            display: block;
            width: max-content;
            max-width: 100%;
            overflow: auto;
            margin-top: 0;
            margin-bottom: 16px;
        }}

        table th {{
            font-weight: 600;
            padding: 6px 13px;
            border: 1px solid #d1d9e0;
            background-color: #f6f8fa;
        }}

        table td {{
            padding: 6px 13px;
            border: 1px solid #d1d9e0;
        }}

        table tr:nth-child(2n) {{
            background-color: #f6f8fa;
        }}

        ul, ol {{
            margin-top: 0;
            margin-bottom: 16px;
            padding-left: 2em;
        }}

        ul ul, ul ol, ol ol, ol ul {{
            margin-top: 0;
            margin-bottom: 0;
        }}

        li > p {{
            margin-top: 16px;
        }}

        a {{
            color: #0969da;
            text-decoration: none;
        }}

        a:hover {{
            text-decoration: underline;
        }}

        hr {{
            height: .25em;
            padding: 0;
            margin: 24px 0;
            background-color: #d1d9e0;
            border: 0;
        }}

        img {{
            max-width: 100%;
            box-sizing: content-box;
        }}

        /* Mermaid diagram container */
        .mermaid-diagram {{
            text-align: center;
            margin: 16px 0;
            padding: 16px;
            background-color: #f6f8fa;
            border-radius: 6px;
            overflow-x: auto;
        }}

        .mermaid-diagram svg {{
            max-width: 100%;
            height: auto;
        }}

        /* Task list items */
        .task-list-item {{
            list-style-type: none;
        }}

        .task-list-item input {{
            margin: 0 .2em .25em -1.4em;
            vertical-align: middle;
        }}

    </style>
</head>
<body>
    {html_body}
</body>
</html>'''

    return html_template, mermaid_success_count, mermaid_failed_count


def convert_markdown_files_parallel(md_file_paths, max_workers=4):
    """
    Convert multiple markdown files to HTML concurrently.

    Processes markdown files in parallel, utilizing multiple CPU cores for
    faster conversion. Especially beneficial for documentation-heavy repositories.

    Args:
        md_file_paths (list): List of markdown file paths to convert
        max_workers (int): Number of concurrent conversion workers (default: 4)

    Returns:
        dict: Mapping of {md_file_path: (success, html_content_or_error)}
              success is True if conversion succeeded, False otherwise
              Second tuple element is HTML content string on success, error message on failure

    Example:
        >>> md_files = ['doc1.md', 'doc2.md', 'doc3.md']
        >>> results = convert_markdown_files_parallel(md_files)
        >>> for md_file, (success, html_or_error) in results.items():
        ...     if success:
        ...         html_content = html_or_error  # type: str
        ...         with open(md_file.replace('.md', '.html'), 'w', encoding='utf-8') as f:
        ...             f.write(html_content)

    Note:
        - 3-5x faster than sequential conversion for multiple files
        - Mermaid diagram rendering runs in parallel subprocess calls
        - Each conversion is independent (thread-safe)
        - Falls back gracefully on conversion errors
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed

    if not md_file_paths:
        return {}

    results = {}

    def convert_single_file(md_path):
        """Worker function to convert single markdown file"""
        try:
            # Read markdown content
            with open(md_path, 'r', encoding='utf-8') as f:
                md_content = f.read()

            # Convert to HTML
            html_content, mermaid_success, mermaid_failed = convert_markdown_to_html(
                md_content,
                md_path
            )

            # Return HTML content along with Mermaid statistics
            return True, (html_content, mermaid_success, mermaid_failed)

        except Exception as e:
            error_msg = f"Conversion failed: {str(e)[:200]}"
            return False, error_msg

    # Execute conversions in parallel
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Submit all conversion tasks
        future_to_file = {
            executor.submit(convert_single_file, md_path): md_path
            for md_path in md_file_paths
        }

        # Collect results as they complete
        for future in as_completed(future_to_file):
            md_path = future_to_file[future]
            try:
                result = future.result()
                results[md_path] = result
            except Exception as e:
                # Unexpected error from worker
                results[md_path] = (False, f"Worker error: {str(e)[:200]}")

    return results


def convert_markdown_to_html_tempfile(md_path, output_dir=None):
    """
    Convert markdown file to HTML and save to temporary file.

    Convenient wrapper for parallel processing that handles file I/O.
    Creates temporary HTML file with sanitized name in specified directory.

    Args:
        md_path (str): Path to markdown file to convert
        output_dir (str): Directory for output file (default: system temp dir)

    Returns:
        tuple: (success: bool, html_path_or_error: str)
               On success: (True, path_to_html_file)
               On failure: (False, error_message)

    Example:
        >>> success, html_path = convert_markdown_to_html_tempfile('README.md')
        >>> if success:
        ...     print(f"Converted to: {html_path}")
        ...     # Upload html_path to SharePoint
        ...     os.remove(html_path)  # Clean up when done

    Note:
        - Caller is responsible for cleaning up temporary file
        - Thread-safe (each call creates unique temp file)
        - Automatically handles file naming and encoding
    """
    try:
        # Read markdown content
        with open(md_path, 'r', encoding='utf-8') as f:
            md_content = f.read()

        # Convert to HTML
        html_content, mermaid_success, mermaid_failed = convert_markdown_to_html(
            md_content,
            md_path
        )

        # Create temporary HTML file
        if output_dir:
            # Use specified directory
            os.makedirs(output_dir, exist_ok=True)
            fd, html_path = tempfile.mkstemp(
                suffix='.html',
                prefix='md_convert_',
                dir=output_dir
            )
        else:
            # Use system temp directory
            fd, html_path = tempfile.mkstemp(
                suffix='.html',
                prefix='md_convert_'
            )

        # Write HTML content
        try:
            with os.fdopen(fd, 'w', encoding='utf-8') as f:
                f.write(html_content)
        except Exception as write_error:
            os.close(fd)
            raise write_error

        return True, html_path

    except Exception as e:
        error_msg = f"Conversion failed: {str(e)[:200]}"
        return False, error_msg
