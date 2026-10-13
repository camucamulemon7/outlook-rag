"""Opt-in, read-only Windows Outlook GetArray/GetNextRow comparison.

No email bodies, embedding API, index writes or identifying values are output.
"""

import argparse
from itertools import islice
import json
import sys
from types import SimpleNamespace

from outlook_rag import app, outlook as outlook_adapter


def probe(folder, max_rows=65, namespace=None):
    table = folder.GetTable("")
    # Reuse one table so both readers observe the same columns and sort order.
    metadata = list(
        islice(
            app.mail_metadata(SimpleNamespace(GetTable=lambda _: table), []), max_rows
        )
    )
    if not metadata:
        raise ValueError("Choose a populated Outlook folder for this check")
    table.MoveToStart()
    for actual, source in metadata:
        row = table.GetNextRow()
        expected = dict(
            entry_id=str(row.Item("EntryID")),
            mail=str(row.Item("MessageClass")).startswith("IPM.Note"),
            received=app.iso(row.Item(outlook_adapter.RECEIVED), utc=True),
            modified=app.iso(row.Item(outlook_adapter.MODIFIED), utc=True),
        )
        if source != "GetTable" or any(
            actual[key] != value for key, value in expected.items()
        ):
            raise ValueError("GetArray and GetNextRow metadata differ")
        if namespace is not None and actual["mail"]:
            item = namespace.GetItemFromID(actual["entry_id"], folder.StoreID)
            for name, tag, key in (("ReceivedTime", outlook_adapter.RECEIVED, "received"),
                                   ("LastModificationTime", outlook_adapter.MODIFIED, "modified")):
                instant = outlook_adapter.item_date(item, name, tag)
                if not app.same_modified(instant, actual[key]):
                    raise ValueError("Item and table UTC properties disagree")
                # Built-in properties are local wall times; their COM tzinfo
                # must not be trusted as UTC. No date or mail ID is printed.
                if not app.same_modified(app.iso(getattr(item, name)), instant):
                    raise ValueError("Local wall time and UTC source disagree")
            item = None
    return dict(
        passed=True,
        rows_checked=len(metadata),
        embedding_api_used=False,
        mail_bodies_read=False,
        outlook_writes=False,
        timezone_comparison="Item UTC/local and table UTC" if namespace is not None else "table readers only",
    )


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--folder", default="inbox")
    parser.add_argument("--max-rows", type=int, default=65)
    args = parser.parse_args()
    if sys.platform != "win32":
        parser.error("This check requires Windows and Outlook Classic")
    if not 1 <= args.max_rows <= 256:
        parser.error("--max-rows must be 1..256")
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    outlook = namespace = folder = None
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        folder = app.resolve_folder(namespace, args.folder)
        result = probe(folder, args.max_rows, namespace)
    except Exception as exc:
        # COM diagnostics can contain account/folder identifiers.
        print(json.dumps(dict(passed=False, error_type=type(exc).__name__)))
        return 1
    finally:
        folder = namespace = outlook = None
        pythoncom.CoUninitialize()
    print(json.dumps(result))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
