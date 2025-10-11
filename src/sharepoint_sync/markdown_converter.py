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
    Sanitize Mermaid diagram code to fix common syntax issues.

    Based on Mermaid.js documentation and common issues, this handles:
    - Self-closing HTML tags (<br/> -> <br>)
    - Double quotes in node labels (break Mermaid syntax)
    - Special characters that break parser (%, ;, #, &, |)
    - Reserved words like "end" (lowercase breaks diagrams)
    - Problematic node prefixes ("o", "x" create unintended edges)
    - HTML tags except <br>
    - Curly braces in comments
    - Diamond/rhombus nodes {}
    - Other node shapes (>], [>, etc.)
    - Edge labels (text between pipes on arrows)
    - Double pipe issues in edge syntax

    Args:
        mermaid_code (str): Raw Mermaid diagram definition

    Returns:
        str: Sanitized Mermaid code safe for mmdc rendering
    """
    sanitized = mermaid_code

    # 1. Replace self-closing <br/> with <br> (Mermaid doesn't support XHTML syntax)
    sanitized = re.sub(r'<br\s*/>', '<br>', sanitized, flags=re.IGNORECASE)

    # 2. Remove other HTML tags except <br>
    # Keep <br> since Mermaid supports it for line breaks
    sanitized = re.sub(r'<(?!br\b)[^>]+>', '', sanitized, flags=re.IGNORECASE)

    # 3. Fix reserved word "end" - it breaks Flowcharts and Sequence diagrams
    # Replace standalone lowercase "end" with "End" in node labels
    # Match patterns like [end], (end), or "end" but not "append", "ending", etc.
    sanitized = re.sub(r'\b(end)\b', 'End', sanitized)

    # 4. Fix double pipes in edge definitions (||) -> (|)
    # Pattern: -->|| or ---|| or ||| should become -->| or ---|
    sanitized = re.sub(r'(\-\->|\-\-\-)\|\|', r'\1|', sanitized)
    sanitized = re.sub(r'\|\|(\w)', r'|\1', sanitized)

    # 5. Escape special characters that break Mermaid syntax
    # Use placeholders first, then replace with entity codes to avoid double-encoding
    def sanitize_content(content):
        """Replace special characters with entity codes using placeholders"""
        # Use temporary placeholders to avoid double-encoding
        content = content.replace('&', '___AMP___')
        content = content.replace('#', '___HASH___')
        content = content.replace('%', '___PERCENT___')
        content = content.replace('|', '___PIPE___')
        content = content.replace('"', '___QUOTE___')

        # Replace placeholders with entity codes
        content = content.replace('___AMP___', '&#38;')
        content = content.replace('___HASH___', '&#35;')
        content = content.replace('___PERCENT___', '&#37;')
        content = content.replace('___PIPE___', '&#124;')
        content = content.replace('___QUOTE___', "'")  # Use single quote instead

        return content

    def sanitize_node_content(match):
        """Replace special characters in square bracket node content"""
        content = match.group(1)
        return f'[{sanitize_content(content)}]'

    # Apply sanitization to content inside square brackets []
    sanitized = re.sub(r'\[([^\]]*)\]', sanitize_node_content, sanitized)

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
    def sanitize_curly_content(match):
        """Replace special characters in curly brace node content"""
        opening_braces = match.group(1)
        content = match.group(2)
        closing_braces = match.group(3)

        return f'{opening_braces}{sanitize_content(content)}{closing_braces}'

    # Match single or double curly braces: {text}, {{text}}
    sanitized = re.sub(r'(\{+)([^{}]+)(\}+)', sanitize_curly_content, sanitized)

    # 8. Handle trapezoid node shapes: [/text\] and [\text/]
    def sanitize_trapezoid_content(match):
        """Replace special characters in trapezoid node content"""
        opening = match.group(1)
        content = match.group(2)
        closing = match.group(3)

        return f'{opening}{sanitize_content(content)}{closing}'

    # Match trapezoid patterns
    sanitized = re.sub(r'(\[/)(.*?)(\\\])', sanitize_trapezoid_content, sanitized)
    sanitized = re.sub(r'(\[\\\x5c)(.*?)(/\])', sanitize_trapezoid_content, sanitized)

    # 9. Handle hexagon node shapes: {{text}}
    # Already handled by curly brace sanitization above

    # 10. Handle edge labels (text between pipes on arrows)
    # Pattern: -->|text| or ---|text|--- etc.
    def sanitize_edge_label(match):
        """Replace special characters in edge labels"""
        prefix = match.group(1)
        label = match.group(2)
        suffix = match.group(3)

        # For edge labels, only escape quotes and special chars that break syntax
        # Don't escape pipes since they're delimiters
        label_sanitized = label.replace('"', "'")
        label_sanitized = label_sanitized.replace('&', '&#38;')
        label_sanitized = label_sanitized.replace('#', '&#35;')
        label_sanitized = label_sanitized.replace('%', '&#37;')

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


def convert_mermaid_to_svg(mermaid_code):
    """
    Convert Mermaid diagram code to SVG using mermaid-cli.

    Uses the mmdc command-line tool installed via npm to render
    Mermaid diagrams as static SVG images.

    Args:
        mermaid_code (str): Mermaid diagram definition

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

        # Run mermaid-cli to convert to SVG
        # Using puppeteer config to work in Docker container
        result = subprocess.run(
            ['mmdc', '-i', mmd_path, '-o', svg_path, '--puppeteerConfigFile', '/usr/src/app/puppeteer-config.json'],
            capture_output=True,
            text=True,
            timeout=30
        )

        if result.returncode == 0 and os.path.exists(svg_path):
            # Read the generated SVG
            with open(svg_path, 'r', encoding='utf-8') as f:
                svg_content = f.read()

            # Clean up temporary files
            os.unlink(mmd_path)
            os.unlink(svg_path)

            return svg_content
        else:
            print(f"[!] Mermaid conversion failed: {result.stderr}")
            # Clean up temp file
            if os.path.exists(mmd_path):
                os.unlink(mmd_path)
            if os.path.exists(svg_path):
                os.unlink(svg_path)
            return None

    except Exception as e:
        print(f"[!] Error converting Mermaid diagram: {e}")
        return None


def convert_markdown_to_html(md_content, filename):
    """
    Convert Markdown content to HTML with Mermaid diagrams rendered as SVG.

    This function:
    1. Parses markdown using Mistune
    2. Finds and converts Mermaid code blocks to inline SVG
    3. Applies GitHub-like styling for SharePoint viewing

    Args:
        md_content (str): Markdown content to convert
        filename (str): Original filename for the HTML title

    Returns:
        str: Complete HTML document with embedded styles and SVGs
    """
    # First, extract and convert all mermaid blocks to placeholder SVGs
    mermaid_pattern = r'```mermaid\n(.*?)\n```'
    mermaid_blocks = []

    def replace_mermaid_with_placeholder(match):
        mermaid_code = match.group(1)
        placeholder = f"<!--MERMAID_PLACEHOLDER_{len(mermaid_blocks)}-->"

        # Convert to SVG
        svg_content = convert_mermaid_to_svg(mermaid_code)
        if svg_content:
            # Clean up the SVG for inline embedding
            # Remove XML declaration if present
            svg_content = re.sub(r'<\?xml[^>]*\?>', '', svg_content)
            svg_content = svg_content.strip()
            mermaid_blocks.append(svg_content)
        else:
            # If conversion failed, keep as code block
            mermaid_blocks.append(f'<pre><code>mermaid\n{mermaid_code}</code></pre>')

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

    return html_template
