"""Pack the tested Paper Maker bundle into download parts for the online installer.

The bundle folder (PaperMaker.exe + engine/ + mineru_home/) is split into three
components so an update only re-downloads what changed::

    app     PaperMaker.exe, READ ME FIRST.txt   (changes with every app build)
    engine  engine/                             (Python + MinerU, rarely changes)
    models  mineru_home/ without the GGUF VLM   (only --tier basic is used)

Each component becomes one compressed tar stream cut into parts of at most
250 MiB (the Cloudflare dashboard uploads files up to 300 MB).  Part names carry
the component's content id, so unchanged components keep their names and need
no re-upload.  ``manifest.json`` lists every part with its size and SHA-256;
PaperMaker_Setup.exe reads it from the bucket.

Usage::

    python installer/pack_release.py "D:/PW Automations/PaperMaker_build/PaperMaker"
        -o "D:/PW Automations/PaperMaker_build/release" [--previous old/manifest.json]
"""
from __future__ import annotations

import argparse
import datetime as dt
import gzip
import hashlib
import io
import json
import lzma
import sys
import tarfile
import time
from pathlib import Path

PART_SIZE = 250 * 1024 * 1024
FORMAT = 1

# name, top-level paths it owns, compression, excluded sub-paths
COMPONENTS = [
    ("app", ["PaperMaker.exe", "READ ME FIRST.txt"], "gz", []),
    ("engine", ["engine"], "xz", []),
    ("models", ["mineru_home"], "gz", ["mineru_home/models/MinerU2.5-Pro-2605-1.2B-GGUF"]),
]


def component_files(bundle: Path, paths: list[str], exclude: list[str]) -> list[Path]:
    files = []
    for p in paths:
        top = bundle / p
        if not top.exists():
            sys.exit(f"missing in bundle: {top}")
        for f in [top] if top.is_file() else sorted(top.rglob("*")):
            rel = f.relative_to(bundle).as_posix()
            if f.is_file() and not any(rel == e or rel.startswith(e + "/") for e in exclude):
                files.append(f)
    return files


def content_id(bundle: Path, files: list[Path]) -> tuple[str, int]:
    """Hash of every path, size and file content: equal id = nothing to re-upload."""
    h, total = hashlib.sha256(), 0
    for f in files:
        fh = hashlib.sha256()
        with f.open("rb") as fp:
            while chunk := fp.read(4 << 20):
                fh.update(chunk)
        size = f.stat().st_size
        total += size
        h.update(f"{f.relative_to(bundle).as_posix()}\0{size}\0{fh.hexdigest()}\n".encode())
    return h.hexdigest(), total


class PartWriter(io.RawIOBase):
    """File object that cuts everything written into numbered parts."""

    def __init__(self, out: Path, stem: str):
        self.out, self.stem, self.parts, self.fp = out, stem, [], None

    def writable(self) -> bool:
        return True

    def _next(self) -> None:
        self._close_part()
        name = f"{self.stem}.{len(self.parts) + 1:03d}"
        self.fp, self.hash, self.size = (self.out / name).open("wb"), hashlib.sha256(), 0
        self.parts.append({"file": name})

    def _close_part(self) -> None:
        if self.fp:
            self.fp.close()
            self.parts[-1].update(size=self.size, sha256=self.hash.hexdigest())
            self.fp = None

    def write(self, b) -> int:
        b, n = memoryview(b), len(b)
        while b:
            if self.fp is None or self.size == PART_SIZE:
                self._next()
            take = b[:PART_SIZE - self.size]
            self.fp.write(take)
            self.hash.update(take)
            self.size += len(take)
            b = b[len(take):]
        return n

    def close(self) -> None:
        self._close_part()
        super().close()


def pack(bundle: Path, out: Path, name: str, files: list[Path], comp: str, cid: str) -> list[dict]:
    writer = PartWriter(out, f"{name}-{cid[:12]}")
    # Python 3.11's "w|gz"/"w|xz" take no level, so compress outside tarfile
    packer = (lzma.LZMAFile(writer, "wb", preset=6) if comp == "xz"
              else gzip.GzipFile(fileobj=writer, mode="wb", compresslevel=6, mtime=0))
    t, done = time.time(), 0
    with tarfile.open(fileobj=packer, mode="w|", format=tarfile.PAX_FORMAT) as tar:
        for i, f in enumerate(files):
            tar.add(f, arcname=f.relative_to(bundle).as_posix(), recursive=False)
            done += f.stat().st_size
            if i % 500 == 0 or i == len(files) - 1:
                print(f"\r  {name}: {done / 2**20:,.0f} MB packed, {time.time() - t:.0f}s",
                      end="", flush=True)
    packer.close()
    writer.close()
    print()
    return writer.parts


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("bundle", type=Path, help="tested bundle folder (contains PaperMaker.exe)")
    ap.add_argument("-o", "--out", type=Path, required=True, help="release folder to write")
    ap.add_argument("--version", default=dt.date.today().strftime("%Y.%m.%d"))
    ap.add_argument("--previous", type=Path, help="manifest.json of the release already in the bucket")
    a = ap.parse_args()

    previous = {}
    if a.previous:
        previous = {c["name"]: c for c in json.loads(a.previous.read_text("utf-8"))["components"]}
    a.out.mkdir(parents=True, exist_ok=True)
    comps, upload = [], []
    for name, paths, comp, exclude in COMPONENTS:
        files = component_files(a.bundle, paths, exclude)
        print(f"{name}: hashing {len(files)} files ...", flush=True)
        cid, unpacked = content_id(a.bundle, files)
        if name in previous and previous[name]["id"] == cid:
            print(f"  unchanged since the previous release, parts reused")
            comps.append(previous[name])
            continue
        parts = pack(a.bundle, a.out, name, files, comp, cid)
        upload += [p["file"] for p in parts]
        comps.append({"name": name, "id": cid, "paths": paths, "compression": comp,
                      "unpacked_size": unpacked, "files": len(files), "parts": parts})
        print(f"  {len(parts)} part(s), {sum(p['size'] for p in parts) / 2**20:,.0f} MB")

    manifest = {"format": FORMAT, "version": a.version, "launch": "PaperMaker.exe",
                "published": dt.datetime.now().isoformat(timespec="seconds"), "components": comps}
    (a.out / "manifest.json").write_text(json.dumps(manifest, indent=1), "utf-8")
    upload.append("manifest.json")
    (a.out / "UPLOAD_THESE.txt").write_text(
        "Upload these files to the bucket root (manifest.json LAST):\n\n" + "\n".join(upload) + "\n",
        "utf-8")
    total = sum(p["size"] for c in comps for p in c["parts"])
    print(f"\nversion {a.version}: download {total / 2**20:,.0f} MB in "
          f"{sum(len(c['parts']) for c in comps)} part(s); to upload: {len(upload)} file(s)")
    print(f"see {a.out / 'UPLOAD_THESE.txt'}")


if __name__ == "__main__":
    main()
