from __future__ import annotations

from typing import Self
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


class EntryData:
    def __init__(
        self,
        name: str,
        stream_id=NOSTREAM,
        obj_type=OBJTY_EMPTY,
        sector=0x00000000,
        size=0,
        parent: Entry = None,
    ):

        # Entry information
        self.name = name[:31].encode("utf-16-le")
        self.stream_id = stream_id
        self.obj_type = obj_type
        self.sector = sector
        self.size = size

        # Directory hierarchy
        self.parent = parent  # Entry of parent directory
        self.child = Directory()  # RBTree of child directory

    def __eq__(self, other: Self) -> bool:
        """
        Regard two `EntryData` instance is the same if the name is identical.

        """

        return self.name == other.name

    def __gt__(self, other: Self) -> bool:
        """
        Define comparison operator as described on MS-CFB documentation.

        Currently, uppercase conversion is achieved by using `upper()` method provided by python. As `upper()` is implemented regarding unicode uppercase conversion rule, it works fine with most of the case.

        However, as MS-CFB has some exception cases for several unicode characters. But this exceptions are negligible. Who will use "LATIN SMALL LETTER A WITH STROKE" for directory name?

        """

        self_name = self.name.decode("utf-16-le").upper()
        other_name = other.name.decode("utf-16-le").upper()

        if len(self_name) != len(other_name):
            return len(self_name) > len(other_name)

        return self_name > other_name


class Entry:
    def __init__(
        self,
        data: EntryData = None,
        color=RED,
        left: Self = None,
        right: Self = None,
        parent: Self = None,
    ):
        # Entry data
        self.data = data

        # RBTree information
        self.color = color
        self.left = left
        self.right = right
        self.parent = parent

    def __bytes__(self) -> bytes:
        """
        Represents directory entry objects as 128-byte bytes string.

        """
        directory_data = struct.pack(
            "<64sHBBIII16sIQQIQ",
            self.data.name,
            len(self.data.name) + 2,
            self.data.obj_type,
            self.color,
            self.left.stream_id(),
            self.right.stream_id(),
            self.data.child.root.stream_id(),
            bytes(16),
            0x00000000,
            0x0000000000000000,
            0x0000000000000000,
            self.data.sector,
            self.data.size,
        )

        return directory_data

    def stream_id(self):
        """
        Get `stream_id` from entry data

        This shortcut is provided since `stream_id` is often used and there are some cases where entry data is `None`.

        """

        if self.data is None:
            return NOSTREAM

        return self.data.stream_id


class Directory:
    def __init__(self):
        """
        Initialize Directory

        Create NIL entry for indicating there's no entry available.
        Set root as NIL.

        """

        self.NIL = Entry(color=BLACK)
        self.root = self.NIL

    def insert(self, data: EntryData):
        new_entry = Entry(data=data, left=self.NIL, right=self.NIL)
        parent = None
        current = self.root

        while current != self.NIL:
            parent = current

            if new_entry.data < current.data:
                current = current.left

            else:
                current = current.right

        new_entry.parent = parent

        # Tree is empty
        if parent is None:
            self.root = new_entry

        elif new_entry.data < parent.data:
            parent.left = new_entry

        else:
            parent.right = new_entry

        new_entry.left = self.NIL
        new_entry.right = self.NIL
        new_entry.color = RED

        self.fix_insert(new_entry)

    def fix_insert(self, entry):
        while entry.parent and entry.parent.color == RED:
            if entry.parent == entry.parent.parent.left:
                uncle = entry.parent.parent.right

                # Case 1: Uncle is red
                if uncle.color == RED:
                    entry.parent.color = BLACK
                    uncle.color = BLACK
                    entry.parent.parent.color = RED
                    entry = entry.parent.parent

                else:
                    # Case 2: entry is right child
                    if entry == entry.parent.right:
                        entry = entry.parent
                        self.left_rotate(entry)

                    # Case 3: entry is left child
                    entry.parent.color = BLACK
                    entry.parent.parent.color = RED
                    self.right_rotate(entry.parent.parent)

            else:
                uncle = entry.parent.parent.left

                # Case 1: Uncle is red
                if uncle.color == RED:
                    entry.parent.color = BLACK
                    uncle.color = BLACK
                    entry.parent.parent.color = RED
                    entry = entry.parent.parent

                else:
                    # Case 2: entry is left child
                    if entry == entry.parent.left:
                        entry = entry.parent
                        self.right_rotate(entry)

                    # Case 3: entry is right child
                    entry.parent.color = BLACK
                    entry.parent.parent.color = RED
                    self.left_rotate(entry.parent.parent)

        self.root.color = BLACK

    def left_rotate(self, entry):
        right_child = entry.right
        entry.right = right_child.left

        if right_child.left != self.NIL:
            right_child.left.parent = entry

        right_child.parent = entry.parent

        # entry is root
        if entry.parent is None:
            self.root = right_child

        elif entry == entry.parent.left:
            entry.parent.left = right_child

        else:
            entry.parent.right = right_child

        right_child.left = entry
        entry.parent = right_child

    def right_rotate(self, entry):
        left_child = entry.left
        entry.left = left_child.right

        if left_child.right != self.NIL:
            left_child.right.parent = entry

        left_child.parent = entry.parent

        # entry is root
        if entry.parent is None:
            self.root = left_child

        elif entry == entry.parent.right:
            entry.parent.right = left_child

        else:
            entry.parent.left = left_child

        left_child.right = entry
        entry.parent = left_child

    def search_name(self, name):
        """
        Search entry by name of entry data.

        """

        # Create dummy entry data for search
        name_entry_data = EntryData(name=name, stream_id=NOSTREAM)

        return self._search_tree_helper(self.root, name_entry_data)

    def _search_tree_helper(self, entry, data):
        if entry == self.NIL or data == entry.data:
            return entry

        if data < entry.data:
            return self._search_tree_helper(entry.left, data)

        return self._search_tree_helper(entry.right, data)

    def traverse(self, entry=None):
        """
        Traverse directory tree

        """

        if entry == None:
            entry = self.root

        if entry != self.NIL:
            yield from self.traverse(entry.left)

            yield entry
            yield from entry.data.child.traverse()

            yield from self.traverse(entry.right)
