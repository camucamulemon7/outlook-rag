"""Opt-in extraction, original-mail preservation and attachment evidence."""
from pathlib import Path
import hashlib
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch
import zipfile

from pypdf import PdfWriter
from pypdf.generic import DictionaryObject, NameObject, DecodedStreamObject
from outlook_rag import app, attachments, outlook


class AttachmentSearch(unittest.TestCase):
    def test_office_and_pdf_extract_in_isolated_worker_and_errors_do_not_stop_mail(self):
        with tempfile.TemporaryDirectory() as root:
            files = []
            for suffix, member in ((".docx", "word/document.xml"), (".xlsx", "xl/sharedStrings.xml"), (".pptx", "ppt/slides/slide1.xml")):
                path = Path(root) / ("fixture" + suffix)
                with zipfile.ZipFile(path, "w") as archive:
                    archive.writestr(member, '<root><t>PCIe attachment specification</t></root>')
                files.append(path)
            pdf = Path(root) / "fixture.pdf"
            writer = PdfWriter()
            page = writer.add_blank_page(612, 792)
            font = DictionaryObject({NameObject("/Type"): NameObject("/Font"), NameObject("/Subtype"): NameObject("/Type1"), NameObject("/BaseFont"): NameObject("/Helvetica")})
            page[NameObject("/Resources")] = DictionaryObject({NameObject("/Font"): DictionaryObject({NameObject("/F1"): writer._add_object(font)})})
            stream = DecodedStreamObject()
            stream.set_data(b"BT /F1 12 Tf 10 700 Td (PCIe attachment specification) Tj ET")
            page[NameObject("/Contents")] = writer._add_object(stream)
            writer.write(pdf)
            files.append(pdf)
            bad = Path(root) / "broken.pdf"
            bad.write_bytes(b"invalid PDF")
            files.append(bad)
            saved = []
            def entry(n):
                source = files[n - 1]
                def save(target):
                    saved.append(Path(target))
                    Path(target).write_bytes(source.read_bytes())
                return SimpleNamespace(FileName=source.name, Size=source.stat().st_size, SaveAsFile=save)
            item = SimpleNamespace(Attachments=SimpleNamespace(Count=len(files), Item=entry))
            found, errors = attachments.read(item, {"index_attachments": True})
            self.assertEqual(len(found), 4)
            self.assertTrue(all("PCIe attachment specification" in entry["text"] for entry in found))
            self.assertEqual(len(errors), 1)
            self.assertTrue(all(not path.exists() for path in saved))
            item.EntryID, item.Subject, item.Body = "fixture", "Specification", "Original message body"
            metadata = dict(entry_id="fixture", received="", modified="", item=item, _attachment_cfg={"index_attachments": True})
            mail = outlook.read_outlook_mail(None, metadata, "store", "archive")
            self.assertEqual(mail["body"], "Original message body")
            self.assertEqual(len(mail["attachment_errors"]), 1)
            metadata["_attachment_cfg"] = {"index_attachments": False}
            with patch.object(attachments, "read", side_effect=AssertionError("Default must not parse")):
                self.assertNotIn("attachments", outlook.read_outlook_mail(None, metadata, "store", "archive"))

    def test_attachment_match_cites_filename_and_identity_is_opt_in(self):
        with tempfile.TemporaryDirectory() as root:
            cfg = dict(model="fixture", dimensions=4, data_dir=root, embedding_url="http://fixture.invalid", index_attachments=True)
            self.assertNotEqual(app.index_identity(cfg), app.index_identity(dict(cfg, index_attachments=False)))
            sql, table = app.open_store(cfg)
            try:
                mail = dict(id=hashlib.sha256(b"attachment mail").hexdigest(), subject="Specification", body="Please see the attachment", attachments=[dict(name="minutes.docx", text="Approved final PCIe timing specification")])
                with patch.object(app, "embed", side_effect=lambda texts, *a, **k: [[1., 0., 0., 0.] for _ in texts]):
                    app.upsert_mail(mail, sql, table, cfg)
                with patch.object(app, "local_config", return_value=cfg):
                    result = app.search("Approved timing", mode="lexical")
                    self.assertEqual(result["items"][0]["attachment_name"], "minutes.docx")
                    self.assertEqual(app.get_mail(mail["id"])["body"], mail["body"])
                    app.remove_mail(mail["id"], sql, table)
                    self.assertEqual(sql.execute("SELECT count(*) FROM chunk_sources").fetchone()[0], 0)
            finally:
                sql.close()
