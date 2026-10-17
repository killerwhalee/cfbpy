from __future__ import annotations

import struct

# Directory entry constants
MAXREGSID = 0xFFFFFFFA  #: (-6) maximum directory entry ID
NOSTREAM = 0xFFFFFFFF  #: (-1) unallocated directory entry

# Object type constants
OBJTY_EMPTY = 0x00  #: empty directory entry
OBJTY_STORAGE = 0x01  #: element is a storage object
OBJTY_STREAM = 0x02  #: element is a stream object
OBJTY_LOCKBYTES = 0x03  #: element is an ILockBytes object
OBJTY_PROPERTY = 0x04  #: element is an IPropertyStorage object
OBJTY_ROOT = 0x05  #: element is a root storage

# Color flags
RED = 0x00
BLACK = 0x01

# Directory entry layout
ENTRY_FORMAT = struct.Struct("<64sHBBIII16sIQQIQ")
ENTRY_SIZE = ENTRY_FORMAT.size  # 128 bytes
MAX_NAME_LENGTH = 31  # UTF-16 code units, excluding the terminating NUL
ILLEGAL_NAME_CHARS = frozenset("/\\:!")

EMPTY_ENTRY = ENTRY_FORMAT.pack(
    b"", 0, OBJTY_EMPTY, RED, NOSTREAM, NOSTREAM, NOSTREAM, bytes(16), 0, 0, 0, 0, 0
)


def sort_key(name: str):
    """
    Sort key for sibling entries as described on MS-CFB documentation.

    Names are compared by length first, then by their uppercase UTF-16 code units. Uppercase conversion uses python's `upper()` per character, keeping characters whose uppercase form is longer than one character (such as "ß") as they are.

    MS-CFB has some exception cases for several unicode characters, but those exceptions are negligible.

    """

    upper = "".join(c.upper() if len(c.upper()) == 1 else c for c in name)

    # Big-endian UTF-16 bytes compare in the same order as UTF-16 code units
    return len(name.encode("utf-16-le")) // 2, upper.encode("utf-16-be")


class Entry:
    """
    Directory entry of a storage or stream.

    Children are kept in a dict keyed by `sort_key`, so lookups and duplicate checks are O(1). The red-black tree is only built when the directory is serialized.

    """

    __slots__ = (
        "name",
        "obj_type",
        "key",
        "children",
        "sector",
        "size",
        "stream_id",
        "color",
        "left",
        "right",
        "child",
    )

    def __init__(self, name: str, obj_type=OBJTY_STREAM, sector=0, size=0):
        encoded = name.encode("utf-16-le")

        if not name or len(encoded) > MAX_NAME_LENGTH * 2:
            raise ValueError(
                f"Entry name must be 1 to {MAX_NAME_LENGTH} UTF-16 code units long: {name!r}"
            )

        if ILLEGAL_NAME_CHARS.intersection(name):
            raise ValueError(f"Entry name must not contain any of '/\\:!': {name!r}")

        # Entry information
        self.name = encoded
        self.obj_type = obj_type
        self.key = sort_key(name)
        self.children: dict[tuple, Entry] = {}
        self.sector = sector
        self.size = size

        # Assigned on serialization
        self.stream_id = NOSTREAM
        self.color = BLACK
        self.left = NOSTREAM
        self.right = NOSTREAM
        self.child = NOSTREAM

    def add(self, entry: Entry) -> Entry:
        """
        Add `entry` as a child of this storage and return it.

        Raises `ValueError` if a sibling with the same name exists, as names are compared case-insensitively.

        """

        if entry.key in self.children:
            raise ValueError(
                f"Duplicate entry name (names are case-insensitive): {entry.name.decode('utf-16-le')!r}"
            )

        self.children[entry.key] = entry
        return entry

    def __bytes__(self) -> bytes:
        """
        Represents directory entry objects as 128-byte bytes string.

        """

        return ENTRY_FORMAT.pack(
            self.name,
            len(self.name) + 2,
            self.obj_type,
            self.color,
            self.left,
            self.right,
            self.child,
            bytes(16),
            0x00000000,
            0x0000000000000000,
            0x0000000000000000,
            self.sector,
            self.size,
        )


def _build_tree(entries: list[Entry]) -> int:
    """
    Link sorted sibling `entries` into a balanced red-black tree.

    Taking the middle entry as root at every level keeps sibling subtree sizes within one of each other, so every leaf lies on the last two levels. Coloring the deepest level red (unless the tree is perfect) and everything else black then satisfies the red-black invariants without any rotation.

    Returns stream ID of the tree root.

    """

    count = len(entries)
    deepest = count.bit_length() - 1
    red_level = deepest if count & (count + 1) else -1

    def link(low, high, level):
        if low >= high:
            return NOSTREAM

        mid = (low + high) // 2
        entry = entries[mid]
        entry.left = link(low, mid, level + 1)
        entry.right = link(mid + 1, high, level + 1)
        entry.color = RED if level == red_level else BLACK

        return entry.stream_id

    return link(0, count, 0)


def serialize(root: Entry, sector_size: int) -> bytes:
    """
    Assign stream IDs, build red-black trees and return directory sector data.

    Stream IDs are assigned in breadth-first order, so siblings get consecutive IDs. Unused entries in the last sector are filled with empty entries.

    """

    entries = [root]
    root.stream_id = 0

    # `entries` grows while iterating, which walks the hierarchy breadth-first
    for entry in entries:
        if not entry.children:
            continue

        siblings = [entry.children[key] for key in sorted(entry.children)]

        for stream_id, sibling in enumerate(siblings, start=len(entries)):
            sibling.stream_id = stream_id

        entries.extend(siblings)
        entry.child = _build_tree(siblings)

    if len(entries) > MAXREGSID:
        raise ValueError("Too many directory entries")

    per_sector = sector_size // ENTRY_SIZE
    padding = -len(entries) % per_sector

    return b"".join(map(bytes, entries)) + EMPTY_ENTRY * padding
