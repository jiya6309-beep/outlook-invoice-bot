"""Writes small but valid Outlook .msg files (Microsoft compound files) for the tests, so no real e-mail has to be shipped.
Every stream is padded to at least 4096 bytes so only regular sectors are needed (the reader supports mini streams too,
and is proven on real Outlook files)."""
import struct

END, FREE, FATSECT = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD


def _filetime(dt):
    import datetime
    return int((dt - datetime.datetime(1601, 1, 1)).total_seconds() * 10_000_000)


class _Entry:
    def __init__(self, name, kind, data=b""):
        self.name, self.kind, self.data, self.children = name, kind, data, []
        self.left = self.right = self.child = FREE
        self.start, self.size = END, 0


def _u16(text):
    return text.encode("utf-16-le")


def _storage(name, streams):
    st = _Entry(name, 1)
    st.children = [_Entry(n, 2, d) for n, d in streams]
    return st


def write_msg(path, subject, received, sender_name="", sender_email="", to="", body="", attachments=()):
    """attachments: list of dicts(filename, data, hidden=False)."""
    props = struct.pack("<8x") + struct.pack("<24x")                       # 32-byte header of the message-level property stream
    props += struct.pack("<IIQ", (0x0E06 << 16) | 0x0040, 6, _filetime(received))
    root_streams = [("__substg1.0_0037001F", _u16(subject)), ("__substg1.0_0C1A001F", _u16(sender_name)),
                    ("__substg1.0_5D01001F", _u16(sender_email)), ("__substg1.0_0E04001F", _u16(to)),
                    ("__substg1.0_1000001F", _u16(body)), ("__properties_version1.0", props)]
    root = _Entry("Root Entry", 5)
    root.children = [_Entry(n, 2, d) for n, d in root_streams]
    for i, a in enumerate(attachments):
        aprops = struct.pack("<8x") + struct.pack("<IIi4x", (0x3705 << 16) | 0x0003, 6, 1)
        aprops += struct.pack("<IIi4x", (0x7FFE << 16) | 0x000B, 6, 1 if a.get("hidden") else 0)
        root.children.append(_storage(f"__attach_version1.0_#{i:08X}", [
            ("__substg1.0_3707001F", _u16(a["filename"])), ("__substg1.0_37010102", a["data"]),
            ("__properties_version1.0", aprops)]))
    # flatten: root first, then every storage and its streams
    flat = [root]
    def add(entry):
        for c in entry.children:
            flat.append(c)
        for c in entry.children:
            if c.kind == 1:
                add(c)
    add(root)
    index = {id(e): i for i, e in enumerate(flat)}
    for e in flat:                                       # link siblings through "right"
        if e.children:
            ids = [index[id(c)] for c in e.children]
            e.child = ids[0]
            for a, b in zip(ids, ids[1:]):
                flat[a].right = b
    # lay the streams out in 512-byte sectors
    sectors, fat = [], []
    def allocate(blob):
        n = max(1, -(-len(blob) // 512))
        first = len(sectors)
        for k in range(n):
            sectors.append(blob[k * 512:(k + 1) * 512].ljust(512, b"\x00"))
            fat.append(first + k + 1 if k < n - 1 else END)
        return first
    for e in flat:
        if e.kind == 2:
            data = e.data
            e.size = len(data)
            e.start = allocate(data.ljust(4096, b"\x00") if len(data) < 4096 else data)
            if len(data) < 4096:
                e.size = 4096
    dir_blob = b""
    for e in flat + [_Entry("", 0)] * ((-len(flat)) % 4):
        name = _u16(e.name)
        entry = bytearray(128)
        entry[0:len(name)] = name
        struct.pack_into("<H", entry, 64, len(name) + 2 if e.name else 0)
        entry[66], entry[67] = e.kind, 1
        struct.pack_into("<III", entry, 68, FREE, e.right if e.kind else FREE, e.child if e.kind else FREE)
        struct.pack_into("<I", entry, 116, e.start if e.kind else 0)
        struct.pack_into("<Q", entry, 120, e.size if e.kind == 2 else 0)
        dir_blob += bytes(entry)
    first_dir = allocate(dir_blob)
    n_fat = 1
    while len(fat) + n_fat > n_fat * 128:
        n_fat += 1
    fat_sectors = list(range(len(sectors), len(sectors) + n_fat))
    fat += [FATSECT] * n_fat
    fat += [FREE] * (n_fat * 128 - len(fat))
    fat_blob = struct.pack(f"<{len(fat)}I", *fat)
    for k in range(n_fat):
        sectors.append(fat_blob[k * 512:(k + 1) * 512])
    header = bytearray(512)
    header[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", header, 24, 0x3E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<IIIIIIII", header, 44, n_fat, first_dir, 0, 4096, END, 0, END, 0)
    difat = fat_sectors + [FREE] * (109 - len(fat_sectors))
    struct.pack_into("<109I", header, 76, *difat)
    with open(path, "wb") as f:
        f.write(bytes(header))
        f.write(b"".join(sectors))
