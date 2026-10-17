"""
Independent, strict reader for compound files written by cfbpy.

It shares no code with the library on purpose: every structure is parsed from
raw bytes and checked against MS-CFB, so a bug in the writer cannot hide behind
the same bug in a reader. Only major version 3 (512-byte sectors) is supported.

"""

import struct

SECTOR_SIZE = 512
MINI_SECTOR_SIZE = 64
MINI_STREAM_CUTOFF_SIZE = 4096
ENTRIES_PER_SECTOR = SECTOR_SIZE // 4

DIFSECT = 0xFFFFFFFC
FATSECT = 0xFFFFFFFD
ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
NOSTREAM = 0xFFFFFFFF

RED, BLACK = 0, 1
EMPTY, STORAGE, STREAM, ROOT = 0, 1, 2, 5


class CFBError(AssertionError):
    pass


def check(condition, message):
    if not condition:
        raise CFBError(message)


def sort_key(name):
    """Directory entry ordering from MS-CFB 2.6.4."""

    upper = "".join(c.upper() if len(c.upper()) == 1 else c for c in name)
    return len(name.encode("utf-16-le")), upper.encode("utf-16-be")


def uint32s(data):
    return list(struct.unpack(f"<{len(data) // 4}I", data))


def read(path):
    """
    Validate the compound file at `path`.

    Returns `{path_tuple: bytes | None}` where `None` marks a storage.

    """

    with open(path, "rb") as fp:
        raw = fp.read()

    # Header
    check(len(raw) >= 512 and (len(raw) - 512) % SECTOR_SIZE == 0, "bad file size")
    (
        signature,
        clsid,
        minor,
        major,
        byte_order,
        sector_shift,
        mini_sector_shift,
        num_dir_sectors,
        num_fat_sectors,
        first_dir_sector,
        _transaction,
        cutoff,
        first_mini_fat_sector,
        num_mini_fat_sectors,
        first_difat_sector,
        num_difat_sectors,
    ) = struct.unpack("<8s16sHHHHH6xIIIIIIIII", raw[:76])

    check(signature == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1", "bad signature")
    check(clsid == bytes(16), "header CLSID must be zero")
    check((minor, major, byte_order) == (0x3E, 3, 0xFFFE), "bad version")
    check((sector_shift, mini_sector_shift) == (9, 6), "bad sector shift")
    check(num_dir_sectors == 0, "num_dir_sectors must be 0 for version 3")
    check(cutoff == MINI_STREAM_CUTOFF_SIZE, "bad mini stream cutoff")

    num_sectors = (len(raw) - 512) // SECTOR_SIZE

    def sector(index):
        check(index < num_sectors, f"sector {index} out of range")
        offset = 512 + index * SECTOR_SIZE
        return raw[offset : offset + SECTOR_SIZE]

    # DIFAT
    difat = uint32s(raw[76:512])
    difat_sectors = []
    next_difat = first_difat_sector
    for _ in range(num_difat_sectors):
        difat_sectors.append(next_difat)
        entries = uint32s(sector(next_difat))
        difat += entries[:-1]
        next_difat = entries[-1]
    check(next_difat == ENDOFCHAIN, "DIFAT chain must end with ENDOFCHAIN")

    fat_sectors = difat[:num_fat_sectors]
    check(FREESECT not in fat_sectors, "DIFAT has fewer entries than FAT sectors")
    check(
        all(entry == FREESECT for entry in difat[num_fat_sectors:]),
        "unused DIFAT entries must be FREESECT",
    )

    # FAT
    fat = []
    for index in fat_sectors:
        fat += uint32s(sector(index))
    check(len(fat) >= num_sectors, "FAT does not cover every sector")
    check(
        all(entry == FREESECT for entry in fat[num_sectors:]),
        "FAT entries past end of file must be FREESECT",
    )
    for index in fat_sectors:
        check(fat[index] == FATSECT, f"FAT sector {index} not marked FATSECT")
    for index in difat_sectors:
        check(fat[index] == DIFSECT, f"DIFAT sector {index} not marked DIFSECT")

    owner = {index: "FAT" for index in fat_sectors}
    owner.update({index: "DIFAT" for index in difat_sectors})

    def chain(table, start, label, limit):
        sectors = []
        current = start
        while current != ENDOFCHAIN:
            check(current < limit, f"{label}: sector {current} out of range")
            sectors.append(current)
            check(len(sectors) <= limit, f"{label}: chain loops")
            current = table[current]
        return sectors

    def claim(sectors, label):
        for index in sectors:
            check(index not in owner, f"{label}: sector {index} already used by {owner.get(index)}")
            owner[index] = label

    def read_chain(start, label):
        sectors = chain(fat, start, label, num_sectors)
        claim(sectors, label)
        return b"".join(sector(index) for index in sectors)

    # Directory
    dir_data = read_chain(first_dir_sector, "directory")
    entries = []
    for offset in range(0, len(dir_data), 128):
        (
            name,
            name_len,
            obj_type,
            color,
            left,
            right,
            child,
            _clsid,
            _state,
            _created,
            _modified,
            start,
            size,
        ) = struct.unpack("<64sHBBIII16sIQQIQ", dir_data[offset : offset + 128])

        if obj_type == EMPTY:
            check(
                (name, name_len, left, right, child) == (bytes(64), 0, NOSTREAM, NOSTREAM, NOSTREAM),
                f"unused directory entry {len(entries)} is not empty",
            )
            entries.append(None)
            continue

        check(name_len % 2 == 0 and 2 <= name_len <= 64, "bad name length")
        check(name[name_len - 2 : name_len] == b"\0\0", "name not NUL terminated")
        check(name[name_len:] == bytes(64 - name_len), "garbage after name")
        entries.append(
            {
                "name": name[: name_len - 2].decode("utf-16-le"),
                "type": obj_type,
                "color": color,
                "left": left,
                "right": right,
                "child": child,
                "start": start,
                "size": size,
            }
        )

    check(entries and entries[0] and entries[0]["type"] == ROOT, "entry 0 must be root")
    check(entries[0]["name"] == "Root Entry", "bad root entry name")

    # Mini stream and mini FAT
    root = entries[0]
    mini_fat_sectors = chain(fat, first_mini_fat_sector, "mini FAT", num_sectors)
    check(len(mini_fat_sectors) == num_mini_fat_sectors, "num_mini_fat_sectors mismatch")
    claim(mini_fat_sectors, "mini FAT")
    mini_fat = []
    for index in mini_fat_sectors:
        mini_fat += uint32s(sector(index))

    mini_stream = read_chain(root["start"], "mini stream") if root["size"] else b""
    if not root["size"]:
        check(root["start"] == ENDOFCHAIN, "empty mini stream must start at ENDOFCHAIN")
    check(root["size"] % MINI_SECTOR_SIZE == 0, "mini stream size not mini-sector aligned")
    check(root["size"] <= len(mini_stream), "mini stream shorter than root size")
    num_mini_sectors = root["size"] // MINI_SECTOR_SIZE
    check(len(mini_fat) >= num_mini_sectors, "mini FAT does not cover mini stream")
    check(
        all(entry == FREESECT for entry in mini_fat[num_mini_sectors:]),
        "unused mini FAT entries must be FREESECT",
    )
    mini_owner = {}

    def read_stream(entry, label):
        size = entry["size"]
        if size == 0:
            check(entry["start"] == ENDOFCHAIN, f"{label}: empty stream must start at ENDOFCHAIN")
            return b""

        if size < MINI_STREAM_CUTOFF_SIZE:
            sectors = chain(mini_fat, entry["start"], label, num_mini_sectors)
            for index in sectors:
                check(index not in mini_owner, f"{label}: mini sector {index} reused")
                mini_owner[index] = label
            data = b"".join(
                mini_stream[index * MINI_SECTOR_SIZE : (index + 1) * MINI_SECTOR_SIZE]
                for index in sectors
            )
            unit = MINI_SECTOR_SIZE
        else:
            data = read_chain(entry["start"], label)
            unit = SECTOR_SIZE

        check(len(data) == -(-size // unit) * unit, f"{label}: chain length does not match size")
        return data[:size]

    # Directory tree
    reached = {0}
    result = {}

    def walk_tree(sid, path, low, high):
        """In-order walk of one storage's tree; returns its black height."""

        if sid == NOSTREAM:
            return 1

        check(sid < len(entries) and entries[sid] is not None, f"bad sid {sid}")
        check(sid not in reached, f"entry {sid} referenced twice")
        reached.add(sid)
        entry = entries[sid]
        key = sort_key(entry["name"])
        check(low is None or low < key, f"tree order broken at {entry['name']!r}")
        check(high is None or key < high, f"tree order broken at {entry['name']!r}")

        for side in ("left", "right"):
            if entry["color"] == RED and entry[side] != NOSTREAM:
                check(entries[entry[side]]["color"] == BLACK, "red node with red child")

        left = walk_tree(entry["left"], path, low, key)
        right = walk_tree(entry["right"], path, key, high)
        check(left == right, f"black height differs under {entry['name']!r}")

        full_path = path + (entry["name"],)
        if entry["type"] == STORAGE:
            check((entry["start"], entry["size"]) == (0, 0), "storage with stream data")
            result[full_path] = None
            walk_storage(entry, full_path)
        elif entry["type"] == STREAM:
            check(entry["child"] == NOSTREAM, "stream with children")
            result[full_path] = read_stream(entry, "/".join(full_path))
        else:
            raise CFBError(f"unexpected object type {entry['type']}")

        return left + (entry["color"] == BLACK)

    def walk_storage(storage, path):
        top = storage["child"]
        if top != NOSTREAM:
            check(top < len(entries) and entries[top] is not None, f"bad sid {top}")
            check(entries[top]["color"] == BLACK, "tree root must be black")
        walk_tree(top, path, None, None)

    walk_storage(root, ())

    used = {sid for sid, entry in enumerate(entries) if entry is not None}
    check(used == reached, f"unreachable directory entries: {sorted(used - reached)}")
    check(len(mini_owner) == num_mini_sectors, "unreferenced mini sectors")
    check(len(owner) == num_sectors, "file contains unreferenced sectors")

    return result
