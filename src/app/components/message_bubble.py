"""
message_bubble.py — Render chat bubbles with citation support.
"""

import re
import streamlit as st


def render_message(role: str, content: str, citations: list = None):
    """
    Render a single chat message bubble.
    role: "user" | "assistant"
    citations: list of {index, title, url, published_date}
    """
    if role == "user":
        st.markdown(
            f"<div class='user-message'>{_escape_html(content)}</div>",
            unsafe_allow_html=True,
        )
    else:
        # Process citation markers [1], [2] into links if citations provided
        rendered = _render_citations_inline(content, citations or [])
        st.markdown(
            f"<div class='ai-message'>{rendered}</div>",
            unsafe_allow_html=True,
        )

        # Render citation list below if present
        if citations:
            _render_citation_list(citations)


def render_thinking():
    """Show animated typing indicator while agent is processing."""
    st.markdown(
        "<div class='thinking'><span></span><span></span><span></span></div>",
        unsafe_allow_html=True,
    )


def render_messages(messages: list):
    """Render full conversation history."""
    for msg in messages:
        render_message(
            role=msg["role"],
            content=msg["content"],
            citations=msg.get("citations", []),
        )


def _render_citations_inline(text: str, citations: list) -> str:
    """Replace [1] markers with linked superscripts if URL available."""
    if not citations:
        return _md_to_safe_html(text)

    citation_map = {str(c["index"]): c for c in citations}

    def replace_marker(match):
        idx = match.group(1)
        if idx in citation_map:
            c = citation_map[idx]
            url = c.get("url", "#")
            return f'<sup><a href="{url}" target="_blank" style="color:#f5a623;text-decoration:none;">[{idx}]</a></sup>'
        return match.group(0)

    text = _md_to_safe_html(text)
    text = re.sub(r'\[(\d+)\]', replace_marker, text)
    return text


def _render_citation_list(citations: list):
    if not citations:
        return
    lines = []
    for c in citations:
        title = c.get("title", f"Source {c['index']}")
        url = c.get("url", "#")
        date = c.get("published_date", "")
        date_str = f" — {date}" if date else ""
        lines.append(f'<a href="{url}" target="_blank">[{c["index"]}] {_escape_html(title)}</a>{date_str}')

    st.markdown(
        f"<div class='citation-list'>{'<br>'.join(lines)}</div>",
        unsafe_allow_html=True,
    )


def _md_to_safe_html(text: str) -> str:
    """Minimal markdown → HTML: bold, italic, code, line breaks."""
    import html
    text = html.escape(text)
    # Bold
    text = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', text)
    # Italic
    text = re.sub(r'\*(.+?)\*', r'<em>\1</em>', text)
    # Inline code
    text = re.sub(r'`(.+?)`', r'<code style="background:#1e1e1e;padding:1px 4px;border-radius:3px;">\1</code>', text)
    # Line breaks
    text = text.replace('\n', '<br>')
    return text


def _escape_html(text: str) -> str:
    import html
    return html.escape(str(text))
