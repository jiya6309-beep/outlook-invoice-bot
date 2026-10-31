"""Reads Outlook .msg files (saved e-mails): subject, sender, recipients, date and attachments.

A .msg file is a Microsoft "compound file". This is a small read-only reader with no extra packages.
It is used only for TESTING (scorecard.py, diagnose.py, run_bot.py --msg-folder): the live bot reads e-mails
straight from Outlook and never needs .msg files."""
import datetime
import os
import struct

FREE, END = 0xFFFFFFFF, 0xFFFFFFFE
MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")
MAX_FILE_MB = 250


class MsgError(Exception):
    """The file is not a readable .msg."""


class _Compound:
    def __init__(self, data):
        d = self.data = data
        if d[:8] != MAGIC:
            raise MsgError("not an Outlook .msg file (compound-file header missing)")
        self.ssz = 1 << struct.unpack_from("<H", d, 30)[0]
        self.mssz = 1 << struct.unpack_from("<H", d, 32)[0]
        n_fat, first_dir, _t, self.cutoff, first_mfat, n_mfat, first_dif, n_dif = struct.unpack_from("<IIIIIIII", d, 44)
        difat = list(struct.unpack_from("<109I", d, 76))
        s = first_dif
        for _ in range(min(n_dif, 4096)):
            vals = struct.unpack("<%dI" % (self.ssz // 4), self._sector(s))
            difat += vals[:-1]
            s = vals[-1]
        self.fat = []
        for s in difat[:n_fat]:
            if s not in (FREE, END):
                self.fat += struct.unpack("<%dI" % (self.ssz // 4), self._sector(s))
        raw = self._chain(first_dir)
        self.entries = []
        for i in range(0, len(raw) - 127, 128):
            e = raw[i:i + 128]
            nlen = struct.unpack_from("<H", e, 64)[0]
            left, right, child = struct.unpack_from("<III", e, 68)
            self.entries.append(dict(name=e[:max(nlen - 2, 0)].decode("utf-16-le", "replace"), type=e[66], left=left,
                                     right=right, child=child, start=struct.unpack_from("<I", e, 116)[0],
                                     size=struct.unpack_from("<Q", e, 120)[0] & 0xFFFFFFFF))
        if not self.entries:
            raise MsgError("empty directory")
        root = self.entries[0]
        self.ministream = self._chain(root["start"], root["size"])
        self.minifat = []
        if n_mfat:
            md = self._chain(first_mfat)
            self.minifat = list(struct.unpack("<%dI" % (len(md) // 4), md))

    def _sector(self, n):
        off = (n + 1) * self.ssz
        sec = self.data[off:off + self.ssz]
        return sec + b"\x00" * (self.ssz - len(sec))           # tolerate a truncated last sector

    def _chain(self, start, size=None):
        out, s, seen = [], start, 0
        while s not in (END, FREE) and s < len(self.fat) and seen <= len(self.fat):
            out.append(self._sector(s))
            s = self.fat[s]
            seen += 1
        blob = b"".join(out)
        return blob[:size] if size is not None else blob

    def _mini_chain(self, start, size):
        out, s, seen = [], start, 0
        while s not in (END, FREE) and s < len(self.minifat) and seen <= len(self.minifat):
            out.append(self.ministream[s * self.mssz:(s + 1) * self.mssz])
            s = self.minifat[s]
            seen += 1
        return b"".join(out)[:size]

    def read(self, idx):
        e = self.entries[idx]
        return self._mini_chain(e["start"], e["size"]) if e["size"] < self.cutoff else self._chain(e["start"], e["size"])

    def children(self, idx):
        out, stack = [], [self.entries[idx]["child"]]
        while stack:
            i = stack.pop()
            if i == FREE or i >= len(self.entries) or i in out:
                continue
            out.append(i)
            stack += [self.entries[i]["left"], self.entries[i]["right"]]
        return out


def _text(raw, kind):
    return raw.decode("utf-16-le", "replace").rstrip("\x00") if kind == "001F" else raw.decode("cp1252", "replace").rstrip("\x00")


def _clean(text):
    return "".join(ch for ch in (text or "") if ch >= " " or ch in "\n\r\t").strip()


def _strings(cf, idx):
    """Variable-length properties ({id: text or bytes}) of a storage."""
    out = {}
    for c in cf.children(idx):
        name = cf.entries[c]["name"]
        if name.startswith("__substg1.0_") and cf.entries[c]["type"] == 2:
            pid, kind = int(name[12:16], 16), name[16:20].upper()
            raw = cf.read(c)
            out[pid] = raw if kind == "0102" else _text(raw, kind)
    return out


def _props(cf, idx, header):
    """Fixed-size properties ({id: int or datetime}) of a storage."""
    out = {}
    for c in cf.children(idx):
        if cf.entries[c]["name"] == "__properties_version1.0":
            blob = cf.read(c)
            for off in range(header, len(blob) - 15, 16):
                tag = struct.unpack_from("<I", blob, off)[0]
                pid, ptype = tag >> 16, tag & 0xFFFF
                raw = blob[off + 8:off + 16]
                if ptype in (0x0003, 0x000B):
                    out[pid] = struct.unpack_from("<i", raw)[0]
                elif ptype == 0x0040:
                    try:
                        out[pid] = datetime.datetime(1601, 1, 1) + datetime.timedelta(microseconds=struct.unpack_from("<Q", raw)[0] // 10)
                    except OverflowError:
                        pass
    return out


def load(path):
    """Parse a .msg. Returns {'subject','sender_name','sender_email','to','received','body','attachments':[...]}.
    Each attachment: {'filename','size','method' (1 = a normal file),'hidden','data'}."""
    if os.path.getsize(path) > MAX_FILE_MB * 1048576:
        raise MsgError(f"larger than {MAX_FILE_MB} MB")
    with open(path, "rb") as f:
        data = f.read()
    try:
        cf = _Compound(data)
        s, pr = _strings(cf, 0), _props(cf, 0, 32)
        msg = dict(subject=_clean(s.get(0x0037)), sender_name=_clean(s.get(0x0C1A)),
                   sender_email=_clean(s.get(0x5D01) or s.get(0x0C1F)), to=_clean(s.get(0x0E04)),
                   received=pr.get(0x0E06) or pr.get(0x0039), body=s.get(0x1000, ""), attachments=[])
        if not msg["body"] and isinstance(s.get(0x1013), bytes):
            msg["body"] = s[0x1013].decode("utf-8", "replace")
        for c in cf.children(0):
            if cf.entries[c]["name"].startswith("__attach_version1.0_"):
                a, ap = _strings(cf, c), _props(cf, c, 8)
                blob = a.get(0x3701) if isinstance(a.get(0x3701), bytes) else None
                msg["attachments"].append(dict(filename=_clean(a.get(0x3707) or a.get(0x3704) or ""),
                                               size=len(blob) if blob else ap.get(0x0E20, 0),
                                               method=ap.get(0x3705, 1), hidden=bool(ap.get(0x7FFE)), data=blob))
        if not (msg["subject"] or msg["sender_email"] or msg["attachments"] or msg["body"]):
            raise MsgError("damaged or truncated .msg (no readable content)")
        return msg
    except MsgError:
        raise
    except Exception as e:                                       # truncated or damaged file
        raise MsgError(f"damaged .msg ({type(e).__name__}: {e})")
