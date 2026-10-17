import array, mmap, os, sys

from . import directory
from .constants import *


class CompoundFileError(ValueError):
    """
    Raised when a file is not a valid compound file.

    """


class DirectoryEntry:
    """
    Directory entry as stored in a compound file.

    """

    __slots__ = ("name", "obj_type", "left", "right", "child", "sector", "size")

    def __init__(self, name, obj_type, left, right, child, sector, size):
        self.name = name
        self.obj_type = obj_type
        self.left = left
        self.right = right
        self.child = child
        self.sector = sector
        self.size = size


def _runs(sectors):
    """
    Group sector numbers into `(first, count)` runs of consecutive sectors.

    """

    first = previous = sectors[0]

    for sector in sectors[1:]:
        if sector != previous + 1:
            yield first, previous - first + 1
            first = sector

        previous = sector

    yield first, previous - first + 1


class CompoundFileReader:
    """
    Read-only access to a compound file of major version 3 or 4.

    The file is memory-mapped, so stream data is copied straight from the page cache without being loaded as a whole. Every sector number and chain is checked, so a malformed file raises `CompoundFileError` instead of looping or reading out of bounds.

    """

    def __init__(self, path):
        self._fp = open(path, "rb")

        try:
            if os.fstat(self._fp.fileno()).st_size < HEADER_SIZE:
                raise CompoundFileError("File is too small to be a compound file")

            self._mmap = mmap.mmap(self._fp.fileno(), 0, access=mmap.ACCESS_READ)
            self._view = memoryview(self._mmap)
            self._load()

        except BaseException:
            self.close()
            raise

    def __enter__(self):
        return self

    def __exit__(self, *exc_info):
        self.close()

    def close(self):
        if view := getattr(self, "_view", None):
            view.release()

        if mapped := getattr(self, "_mmap", None):
            try:
                mapped.close()

            # Slices still referenced elsewhere (e.g. by a traceback) keep the map open until collected
            except BufferError:
                pass

        self._fp.close()

    def _load(self):
        (
            signature,
            _clsid,
            _minor_version,
            major_version,
            byte_order,
            sector_shift,
            mini_sector_shift,
            _num_dir_sectors,
            num_fat_sectors,
            first_dir_sector,
            _transaction_signature_number,
            mini_stream_cutoff_size,
            first_mini_fat_sector,
            _num_mini_fat_sectors,
            first_difat_sector,
            num_difat_sectors,
            *header_difat,
        ) = HEADER_FORMAT.unpack_from(self._view)

        if signature != HEADER_SIGNATURE:
            raise CompoundFileError("Not a compound file: bad signature")

        if byte_order != BYTE_ORDER:
            raise CompoundFileError("Bad byte order mark")

        if (major_version, sector_shift) not in ((3, 9), (4, 12)):
            raise CompoundFileError(
                f"Unsupported major version {major_version} with sector shift {sector_shift}"
            )

        if mini_sector_shift != MINI_SECTOR_SHIFT or mini_stream_cutoff_size != MINI_STREAM_CUTOFF_SIZE:
            raise CompoundFileError("Unsupported mini sector shift or mini stream cutoff size")

        self.major_version = major_version
        self.sector_shift = sector_shift
        self.sector_size = 1 << sector_shift

        # Sector N starts right after the header sector; a partial last sector is allowed
        self.num_sectors = -(-(len(self._view) - self.sector_size) // self.sector_size)

        # Locate FAT sectors through header DIFAT and DIFAT sector chain
        if num_fat_sectors > self.num_sectors or num_difat_sectors > self.num_sectors:
            raise CompoundFileError("Too many FAT or DIFAT sectors for file size")

        fat_sectors = header_difat
        difat_sector = first_difat_sector

        for _ in range(num_difat_sectors):
            if difat_sector in (ENDOFCHAIN, FREESECT):
                break

            entries = self._table(self._sector(difat_sector))
            fat_sectors.extend(entries[:-1])
            difat_sector = entries[-1]

        if len(fat_sectors) < num_fat_sectors:
            raise CompoundFileError("DIFAT has fewer entries than FAT sectors")

        self.fat = self._table(b"".join(self._sector(sector) for sector in fat_sectors[:num_fat_sectors]))

        # Directory
        directory_data = self._read_chain(first_dir_sector)
        self.entries = [self._parse_entry(*fields) for fields in directory.ENTRY_FORMAT.iter_unpack(directory_data)]

        root = self.entries[0] if self.entries else None
        if root is None or root.obj_type != directory.OBJTY_ROOT:
            raise CompoundFileError("First directory entry is not a root entry")

        # Mini stream is stored in regular sectors starting from root entry
        num_mini_stream_sectors = -(-root.size // self.sector_size)
        self._mini_stream_sectors = self._chain(self.fat, root.sector, num_mini_stream_sectors, self.num_sectors)
        self.num_mini_sectors = -(-root.size >> MINI_SECTOR_SHIFT)

        if self.num_mini_sectors:
            self.mini_fat = self._table(self._read_chain(first_mini_fat_sector))
        else:
            self.mini_fat = array.array(UINT32)

    def _sector(self, sector, full=True):
        """
        Return memoryview of sector data.

        """

        if sector >= self.num_sectors:
            raise CompoundFileError(f"Sector number out of range: {sector:#x}")

        offset = (sector + 1) << self.sector_shift
        data = self._view[offset : offset + self.sector_size]

        if full and len(data) < self.sector_size:
            raise CompoundFileError("File is truncated")

        return data

    @staticmethod
    def _table(data) -> array.array:
        """
        Unpack little-endian 32-bit sector numbers.

        """

        # `frombytes()` accepts any bytes-like object, unlike the constructor which iterates a memoryview
        table = array.array(UINT32)
        table.frombytes(data)

        if sys.byteorder == "big":
            table.byteswap()

        return table

    @staticmethod
    def _chain(table, start, count, limit) -> array.array:
        """
        Follow sector chain from `start`.

        Returns first `count` sectors of the chain, or the whole chain if `count` is `None`. `limit` is the number of valid sector numbers.

        """

        sectors = array.array(UINT32)
        current = start

        while current != ENDOFCHAIN and (count is None or len(sectors) < count):
            if current >= limit or current >= len(table):
                raise CompoundFileError(f"Invalid sector number in chain: {current:#x}")

            sectors.append(current)

            # A chain cannot be longer than the number of sectors without looping
            if len(sectors) > limit:
                raise CompoundFileError("Sector chain contains a loop")

            current = table[current]

        if count is not None and len(sectors) < count:
            raise CompoundFileError("Sector chain is shorter than stream size")

        return sectors

    def _read_chain(self, start) -> bytes:
        """
        Read every sector in the FAT chain from `start`.

        """

        sectors = self._chain(self.fat, start, None, self.num_sectors)

        return b"".join(self._sector(sector) for sector in sectors)

    def _parse_entry(self, name, name_length, obj_type, _color, left, right, child, _clsid, _state, _created, _modified, sector, size):
        # Unused entries and object types other than storage/stream are ignored
        if obj_type not in (directory.OBJTY_ROOT, directory.OBJTY_STORAGE, directory.OBJTY_STREAM):
            return None

        if name_length < 2 or name_length > 64 or name_length % 2:
            raise CompoundFileError(f"Invalid directory entry name length: {name_length}")

        try:
            name = name[: name_length - 2].decode("utf-16-le")
        except UnicodeDecodeError as e:
            raise CompoundFileError("Invalid directory entry name") from e

        # High 32 bits of stream size must be ignored for major version 3
        if self.major_version == 3:
            size &= 0xFFFFFFFF

        return DirectoryEntry(name, obj_type, left, right, child, sector, size)

    def _entry(self, stream_id) -> DirectoryEntry:
        entry = self.entries[stream_id] if stream_id < len(self.entries) else None

        if entry is None or entry.obj_type == directory.OBJTY_ROOT:
            raise CompoundFileError(f"Invalid directory entry reference: {stream_id:#x}")

        return entry

    def walk(self):
        """
        Yield `(path, entry)` for every storage and stream, where `path` is a tuple of names.

        Storages are yielded before their children. Siblings are not yielded in any particular order.

        """

        visited = set()
        pending = [((), self.entries[0].child)]

        # Walk red-black trees iteratively, as trees written by other tools may be degenerate
        while pending:
            path, stream_id = pending.pop()

            if stream_id == directory.NOSTREAM:
                continue

            if stream_id in visited:
                raise CompoundFileError("Directory entry referenced more than once")

            visited.add(stream_id)
            entry = self._entry(stream_id)
            pending.append((path, entry.left))
            pending.append((path, entry.right))

            entry_path = path + (entry.name,)
            yield entry_path, entry

            if entry.obj_type == directory.OBJTY_STORAGE:
                pending.append((entry_path, entry.child))

    def _mini_sector(self, sector):
        """
        Return memoryview of mini sector data.

        """

        offset = sector << MINI_SECTOR_SHIFT
        regular = self._mini_stream_sectors[offset >> self.sector_shift]
        start = offset & (self.sector_size - 1)

        return self._sector(regular, full=False)[start : start + (1 << MINI_SECTOR_SHIFT)]

    def copy_stream(self, entry: DirectoryEntry, fp):
        """
        Write data of stream `entry` to file object `fp`.

        """

        size = entry.size
        if not size:
            return

        if size < MINI_STREAM_CUTOFF_SIZE:
            count = -(-size >> MINI_SECTOR_SHIFT)
            sectors = self._chain(self.mini_fat, entry.sector, count, self.num_mini_sectors)
            data = b"".join(self._mini_sector(sector) for sector in sectors)[:size]

            if len(data) < size:
                raise CompoundFileError("File is truncated")

            fp.write(data)
            return

        count = -(-size >> self.sector_shift)
        sectors = self._chain(self.fat, entry.sector, count, self.num_sectors)
        remaining = size

        # Copy each run of consecutive sectors at once
        for first, length in _runs(sectors):
            offset = (first + 1) << self.sector_shift
            length = min(length << self.sector_shift, remaining)
            data = self._view[offset : offset + length]

            if len(data) < length:
                raise CompoundFileError("File is truncated")

            fp.write(data)
            remaining -= length
