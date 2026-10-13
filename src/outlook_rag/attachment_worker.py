"""Bounded, isolated extraction; never imports the MCP server or credentials."""
from __future__ import annotations
import json
import os
from pathlib import Path
import sys
import zipfile

MAX_TEXT = 100000


def constrain() -> object:
    limit = 1024 * 1024 * 1024
    if os.name == "nt":
        import win32api
        import win32job
        job = win32job.CreateJobObject(None, "")
        info = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
        info["BasicLimitInformation"]["LimitFlags"] = win32job.JOB_OBJECT_LIMIT_PROCESS_MEMORY
        info["ProcessMemoryLimit"] = limit
        win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, info)
        win32job.AssignProcessToJobObject(job, win32api.GetCurrentProcess())
        return job
    import resource
    if sys.platform != "darwin":
        resource.setrlimit(resource.RLIMIT_AS, (limit, limit))
    resource.setrlimit(resource.RLIMIT_CPU, (8, 8))
    return None


def extract(path: Path) -> str:
    if not 0 < path.stat().st_size <= 10 * 1024 * 1024:
        raise ValueError("Attachment size limit")
    suffix = path.suffix.casefold()
    if suffix in {".txt", ".csv", ".md"}:
        raw = path.read_bytes()
        encodings = ("utf-16",) if raw.startswith((b"\xff\xfe", b"\xfe\xff")) else ("utf-8-sig", "cp932")
        for encoding in encodings:
            try:
                return raw.decode(encoding)[:MAX_TEXT]
            except UnicodeError:
                pass
        raise ValueError("Unsupported text encoding")
    if suffix == ".pdf":
        from pypdf import PdfReader
        reader = PdfReader(path, strict=True)
        parts: list[str] = []
        size = 0
        for page in reader.pages[:50]:
            value = (page.extract_text() or "")[:MAX_TEXT - size]
            parts.append(value)
            size += len(value)
            if size >= MAX_TEXT:
                break
        return "\n".join(parts)[:MAX_TEXT]
    if suffix not in {".docx", ".xlsx", ".pptx"}:
        raise ValueError("Unsupported attachment format")
    from defusedxml.ElementTree import fromstring
    with zipfile.ZipFile(path) as archive:
        entries = archive.infolist()
        if len(entries) > 1000 or sum(e.file_size for e in entries) > 20 * 1024 * 1024:
            raise ValueError("Office expansion limit")
        prefixes = {".docx": ("word/document.xml", "word/header", "word/footer"),
                    ".xlsx": ("xl/sharedStrings.xml", "xl/worksheets/"), ".pptx": ("ppt/slides/slide", "ppt/notesSlides/notesSlide")}[suffix]
        parts, size = [], 0
        for entry in sorted(entries, key=lambda e: e.filename):
            if not entry.filename.endswith(".xml") or not entry.filename.startswith(prefixes):
                continue
            document = fromstring(archive.read(entry))
            for element in document.iter():
                if element.tag.rsplit("}", 1)[-1] in {"t", "v"} and element.text:
                    value = element.text[:MAX_TEXT - size]
                    parts.append(value)
                    size += len(value)
                    if size >= MAX_TEXT:
                        return "\n".join(parts)[:MAX_TEXT]
        return "\n".join(parts)[:MAX_TEXT]


def main() -> None:
    guard = constrain()
    try:
        text = extract(Path(sys.argv[1]))
        print(json.dumps({"text": text}, ensure_ascii=False), flush=True)
    except Exception as exc:
        print(json.dumps({"error_type": type(exc).__name__}), flush=True)
        raise SystemExit(1)
    # Retain the Windows memory-limit handle until process exit.
    _ = guard


if __name__ == "__main__":
    main()
