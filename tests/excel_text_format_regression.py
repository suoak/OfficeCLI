#!/usr/bin/env python3
# SPDX-License-Identifier: Apache-2.0
"""Check persisted XLSX types through the real CLI (Python standard library only).

Build officecli, then run this file with --officecli /path/to/officecli
(or --officecli /path/to/dotnet /path/to/officecli.dll).
Each case uses its own copy; close flushes the resident before XML inspection.
"""
import argparse
from contextlib import ExitStack
import os
import posixpath
from pathlib import Path
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
Q = lambda name: "{" + NS + "}" + name


def read_package(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def sheet_part(package):
    book = ET.fromstring(package["xl/workbook.xml"])
    sheet = next(node for node in book.find(Q("sheets")) if node.get("name") == "Sheet1")
    rid = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
    rels = ET.fromstring(package["xl/_rels/workbook.xml.rels"])
    target = next(node.get("Target") for node in rels if node.get("Id") == rid)
    part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
    assert part in package, (rid, part)
    return part


def xml_bytes(root, namespace):
    # System.IO.Packaging expects OPC roots to use their default namespace.
    ET.register_namespace("", namespace)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def seed(path):
    package = read_package(path)
    sheet = ET.Element(Q("worksheet"))
    data = ET.SubElement(sheet, Q("sheetData"))
    kinds = [(None, "123"), ("n", "123"), ("b", "1"), ("e", "#DIV/0!"),
             (None, "3"), ("s", "0"), ("inlineStr", "inline text")]
    for index, (kind, value) in enumerate(kinds, 1):
        row = ET.SubElement(data, Q("row"), {"r": str(index)})
        cell = ET.SubElement(row, Q("c"), {"r": "A" + str(index)})
        if kind is not None:
            cell.set("t", kind)
        if index == 5:
            ET.SubElement(cell, Q("f")).text = "1+2"
        if kind == "inlineStr":
            ET.SubElement(ET.SubElement(cell, Q("is")), Q("t")).text = value
        else:
            ET.SubElement(cell, Q("v")).text = value
    package[sheet_part(package)] = xml_bytes(sheet, NS)
    strings = ET.Element(Q("sst"), {"count": "1", "uniqueCount": "1"})
    ET.SubElement(ET.SubElement(strings, Q("si")), Q("t")).text = "shared text"
    package["xl/sharedStrings.xml"] = xml_bytes(strings, NS)
    rels = ET.fromstring(package["xl/_rels/workbook.xml.rels"])
    reltype = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    if not any(node.get("Type") == reltype for node in rels):
        existing = {node.get("Id") for node in rels}
        rid = "rIdTextRegression"
        assert rid not in existing
        ET.SubElement(rels, "{" + REL + "}Relationship", {"Id": rid, "Type": reltype, "Target": "sharedStrings.xml"})
    package["xl/_rels/workbook.xml.rels"] = xml_bytes(rels, REL)
    types = ET.fromstring(package["[Content_Types].xml"])
    if not any(node.get("PartName") == "/xl/sharedStrings.xml" for node in types):
        ET.SubElement(types, "{" + CT + "}Override", {"PartName": "/xl/sharedStrings.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"})
    package["[Content_Types].xml"] = xml_bytes(types, CT)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in package.items():
            archive.writestr(name, content)


def cells(path):
    package = read_package(path)
    sheet = ET.fromstring(package[sheet_part(package)])
    return package, {cell.get("r"): cell for cell in sheet.iter(Q("c"))}


def value_signature(cell):
    inline = cell.find(Q("is"))
    return (cell.get("t"), cell.findtext(Q("v")), cell.findtext(Q("f")),
            None if inline is None else "".join(inline.itertext()))


def assert_text_format(package, cell):
    styles = ET.fromstring(package["xl/styles.xml"])
    xfs = list(styles.find(Q("cellXfs")))
    fmt = xfs[int(cell.get("s", "0"))].get("numFmtId", "0")
    custom = styles.find(Q("numFmts"))
    assert fmt == "49" or (custom is not None and any(
        node.get("numFmtId") == fmt and node.get("formatCode") == "@" for node in custom)), fmt


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--officecli", nargs="+", required=True, help="built CLI executable, or dotnet followed by DLL")
    parser.add_argument("--action", choices=("add", "set"), help="isolate one mutation path for negative controls")
    options = parser.parse_args()
    cli = options.officecli
    actions = (options.action,) if options.action else ("add", "set")
    env = dict(os.environ, OFFICECLI_RESIDENT_FLUSH="each")
    registered = set()
    def close_for_cleanup(path):
        subprocess.run([*cli, "close", str(path)], env=env, capture_output=True, text=True, timeout=60, check=True)
    def run(*args):
        if args[0] in ("create", "add", "set"):
            path = Path(args[1])
            if path not in registered:
                cleanup.callback(close_for_cleanup, path)
                registered.add(path)
        result = subprocess.run([*cli, *map(str, args)], env=env, capture_output=True, text=True, timeout=60)
        assert result.returncode == 0, (args, result.stdout, result.stderr)
    def change(action, path, address, props):
        args = [action, path, "/Sheet1/" + address]
        if action == "add":
            args += ["--type", "cell"]
        for prop in props:
            args += ["--prop", prop]
        run(*args)
    checks = 0
    with tempfile.TemporaryDirectory(prefix="officecli-text-format-") as directory, ExitStack() as cleanup:
        root = Path(directory)
        base = root / "seed.xlsx"
        run("create", base)
        run("close", base)
        book = ET.fromstring(read_package(base)["xl/workbook.xml"])
        sheets = book.find(Q("sheets"))
        if sheets is None or not any(node.get("name") == "Sheet1" for node in sheets):
            run("add", base, "/", "--type", "sheet", "--prop", "name=Sheet1")
            run("close", base)
        seed(base)
        _, before = cells(base)
        for action in actions:
            path = root / (action + "-format.xlsx")
            shutil.copyfile(base, path)
            for address in before:
                change(action, path, address, ["numberformat=@"])
            run("close", path)
            package, after = cells(path)
            for address, original in before.items():
                assert value_signature(after[address]) == value_signature(original), (action, address, value_signature(after[address]))
                assert_text_format(package, after[address])
                checks += 1
            shared = ET.fromstring(package["xl/sharedStrings.xml"])
            assert "".join(list(shared)[int(after["A6"].findtext(Q("v")))].itertext()) == "shared text"
            for alias in ("value", "text"):
                path = root / (action + "-" + alias + ".xlsx")
                shutil.copyfile(base, path)
                change(action, path, "A1", [alias + "=00456", "numberformat=@"])
                run("close", path)
                package, after = cells(path)
                assert after["A1"].get("t") == "str" and after["A1"].findtext(Q("v")) == "00456", (action, alias, value_signature(after["A1"]))
                assert_text_format(package, after["A1"])
                checks += 1
        # The Set path explicitly inherits an existing text number format.
        for alias in (("value", "text") if "set" in actions else ()):
            path = root / ("set-inherit-" + alias + ".xlsx")
            shutil.copyfile(base, path)
            change("set", path, "A1", ["numberformat=@"])
            change("set", path, "A1", [alias + "=00456"])
            run("close", path)
            package, after = cells(path)
            assert after["A1"].get("t") == "str" and after["A1"].findtext(Q("v")) == "00456", (alias, value_signature(after["A1"]))
            assert_text_format(package, after["A1"])
            checks += 1
    print(f"PASS: {checks} persisted OpenXML cell/value/type/format cases")


if __name__ == "__main__":
    if not __debug__:
        raise SystemExit("Run this regression script without Python -O")
    main()
