"""Conservative plain-text cleaning, chunking and Japanese tokenization."""

from __future__ import annotations

import re
import threading
from typing import Any
from urllib.parse import urlsplit

_LOCK = threading.RLock()
_DICTIONARY: Any = None
_LOCAL = threading.local()


def _get_tokenizer() -> tuple[Any, Any]:
    global _DICTIONARY
    if _DICTIONARY is None:
        with _LOCK:
            if _DICTIONARY is None:
                from sudachipy import dictionary

                _DICTIONARY = dictionary.Dictionary()
    tokenizer_obj = getattr(_LOCAL, "tokenizer", None)
    if tokenizer_obj is None:
        from sudachipy import tokenizer

        tokenizer_obj = (_DICTIONARY.tokenizer(), tokenizer.Tokenizer.SplitMode.A)
        _LOCAL.tokenizer = tokenizer_obj
    return tokenizer_obj


def tokens(text: str) -> str:
    engine, mode = _get_tokenizer()
    words = [
        word.normalized_form().lower()
        for word in engine.tokenize(text, mode)
        if not word.part_of_speech()[0] in ("補助記号", "空白", "助詞", "助動詞")
    ]
    return " ".join(words)


def safe_text(value: str) -> str:
    # Outlook COM can expose lone UTF-16 surrogates. Combine valid pairs and
    # replace only invalid units before UTF-8 hashing, HTTP or SQLite writes.
    if re.search(r"[\ud800-\udfff]", value):
        return value.encode("utf-16-le", "surrogatepass").decode("utf-16-le", "replace")
    return value


def clean_body(body: str, version: int = 2) -> str:
    # Keep original body in SQLite. Conservative quote removal for embedding only.
    body = safe_text(body).replace("\r\n", "\n").replace("\r", "\n")
    if version == 3:
        cleaned = clean_body(body, 2)
        significant = [line for line in cleaned.splitlines() if line.strip() and not re.match(
            r"(?i)^\s*(?:From|To|Cc|Sent|Date|Subject|差出人|宛先|送信日時|件名)\s*[:：]", line
        )]
        # Quote-only forwards have no authored reply; keep the only copy of
        # their facts instead of returning an empty or headers-only document.
        return cleaned if significant else re.sub(r"(?m)^\s*> ?", "", body).strip()
    if version not in (1, 2):
        raise ValueError("text_cleaning_version must be 1, 2 or 3")
    source = body.splitlines()
    lines: list[str] = []
    for position, line in enumerate(source):
        if re.match(
            r"^\s*(?:-{3,}\s*Original Message\s*-{3,}|On .{5,200}wrote:)", line, re.I
        ):
            break
        if version == 2:
            # Outlook reply headers require multiple matching labels, not just 'From:'.
            if any(previous.strip() for previous in lines) and re.match(
                r"^\s*(?:From|差出人)\s*[:：]", line, re.I
            ):
                header = "\n".join(source[position : position + 8])
                if re.search(
                    r"(?im)^\s*(?:Sent|送信日時|送信日)\s*[:：]", header
                ) and re.search(r"(?im)^\s*(?:Subject|件名)\s*[:：]", header):
                    break
            # A conventional short signature at the end; ordinary separators remain.
            if line == "-- " and lines and len("\n".join(source[position:])) <= 800:
                break
            line = re.sub(
                r"[\u00ad\u034f\u200b\u200c\u200e\u200f\u2060\ufeff]", "", line
            )
            if re.match(
                r"^\s*(?:配信停止(?:はこちら|をご希望|のお手続き)|購読解除(?:はこちら|のお手続き)|unsubscribe\b)",
                line,
                re.I,
            ):
                continue
            if re.fullmatch(
                r"\s*<?https?://\S+\.(?:png|gif|jpe?g|webp)(?:\?[^\s>]*)?>?\s*",
                line,
                re.I,
            ):
                continue

            def compact_url(match: re.Match[str]) -> str:
                value = match.group(0)
                if len(value) <= 160:
                    return value
                try:
                    parsed = urlsplit(value)
                except ValueError:
                    return value
                # Long tracking links have little semantic value; retain source host.
                return f"{parsed.scheme}://{parsed.netloc}/"

            line = re.sub(r"https?://[^\s<>]+", compact_url, line)
        if not line.lstrip().startswith(">"):
            lines.append(line.rstrip())
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()


def chunks(
    subject: str,
    body: str,
    size: int = 1600,
    overlap: int = 200,
    *,
    cleaning_version: int = 2,
) -> list[str]:
    subject = safe_text(subject)
    text = clean_body(body, cleaning_version)
    if cleaning_version == 3 and re.match(r"(?i)^\s*(?:fw|fwd|転送)\s*[:：]", subject):
        text = re.sub(r"(?m)^\s*> ?", "", safe_text(body).replace("\r\n", "\n").replace("\r", "\n")).strip()
    if not text:
        return [f"Subject: {subject}"]
    if cleaning_version == 3:
        return [f"Subject: {subject}\n\n{part}" for part in paragraph_chunks(text, size, overlap)]
    output, start = [], 0
    while start < len(text):
        end = min(start + size, len(text))
        if end < len(text):
            boundary = text.rfind("\n", start + size // 2, end)
            if boundary > start:
                end = boundary
        output.append(f"Subject: {subject}\n\n{text[start:end]}")
        if end == len(text):
            break
        start = max(start + 1, end - overlap)
    return output


def paragraph_chunks(text: str, size: int, overlap: int) -> list[str]:
    """Pack intact paragraphs; split only oversized ones, at sentence boundaries."""
    atoms = []
    for paragraph in re.split(r"\n\s*\n", text):
        start = 0
        while start < len(paragraph):
            end = min(start + size, len(paragraph))
            if end < len(paragraph):
                boundaries = list(re.finditer(r"[。！？!?]|[.]\s|\n|\s+", paragraph[start + size // 2:end]))
                if boundaries:
                    end = start + size // 2 + boundaries[-1].end()
            atoms.append(paragraph[start:end])
            start = end
    output: list[str] = []
    current: list[str] = []
    for atom in atoms:
        if current and len("\n\n".join([*current, atom])) > size:
            output.append("\n\n".join(current))
            suffix: list[str] = []
            for block in reversed(current):
                if len("\n\n".join([block, *suffix])) > overlap:
                    break
                suffix.insert(0, block)
            current = suffix if len("\n\n".join([*suffix, atom])) <= size else []
        current.append(atom)
    if current:
        output.append("\n\n".join(current))
    return output
