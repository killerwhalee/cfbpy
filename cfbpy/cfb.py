import array, os, sys

from . import directory, reader
from .constants import *
from .reader import CompoundFileError

SECTOR_SIZE = 1 << SECTOR_SHIFT
MINI_SECTOR_SIZE = 1 << MINI_SECTOR_SHIFT
ENTRIES_PER_SECTOR = SECTOR_SIZE // 4  # FAT/MINIFAT/DIFAT entries per sector

# Chunk size for copying large streams
COPY_CHUNK_SIZE = 1 << 20


def _pack_sectors(entries) -> bytes:
    """
    Pack 32-bit sector numbers as little-endian, padded with FREESECT to a whole sector.

    """

    table = array.array(UINT32, entries)
    table.extend([FREESECT] * (-len(table) % ENTRIES_PER_SECTOR))

    if sys.byteorder == "big":
        table.byteswap()

    return table.tobytes()


def _extend_chain(table: array.array, count: int) -> int:
    """
    Allocate `count` consecutive sectors at the end of `table` as one chain.

    Returns first sector number of the chain, or ENDOFCHAIN if `count` is zero.

    """

    if not count:
        return ENDOFCHAIN

    start = len(table)
    table.extend(range(start + 1, start + count))
    table.append(ENDOFCHAIN)

    return start


class _MiniStream:
    """
    Mini stream and its MINIFAT, kept in memory as every stream is smaller than the cutoff size.

    """

    def __init__(self) -> None:
        self.data = bytearray()
        self.fat = array.array(UINT32)

    def write(self, data: bytes) -> int:
        """
        Append data to the mini stream and return its first mini sector number.

        """

        count = -(-len(data) // MINI_SECTOR_SIZE)
        self.data += data
        self.data += bytes(count * MINI_SECTOR_SIZE - len(data))

        return _extend_chain(self.fat, count)


class _SectorWriter:
    """
    Write sectors sequentially to a file object while building the FAT.

    Every stream is written as one run of consecutive sectors, so FAT chains are simple ranges and stream data never has to be held in memory as a whole.

    """

    def __init__(self, fp) -> None:
        self.fp = fp
        self.fat = array.array(UINT32)

        # Reserve space for the header, which is written by `close()`
        fp.write(bytes(HEADER_SIZE))

    def _pad(self, size: int) -> int:
        """
        Pad data written so far to a whole sector and return the number of sectors used.

        """

        count = -(-size // SECTOR_SIZE)
        self.fp.write(bytes(count * SECTOR_SIZE - size))

        return count

    def write(self, data: bytes) -> int:
        """
        Write data to new sectors and return its first sector number.

        """

        self.fp.write(data)

        return _extend_chain(self.fat, self._pad(len(data)))

    def copy(self, head: bytes, src) -> tuple[int, int]:
        """
        Write `head` followed by the rest of file object `src` to new sectors.

        Returns first sector number and size of the stream.

        """

        self.fp.write(head)
        size = len(head)

        while chunk := src.read(COPY_CHUNK_SIZE):
            self.fp.write(chunk)
            size += len(chunk)

        return _extend_chain(self.fat, self._pad(size)), size

    def close(self, **header) -> None:
        """
        Write FAT and DIFAT sectors, then the header.

        """

        num_data_sectors = len(self.fat)

        # FAT and DIFAT sectors need FAT entries themselves, so grow both until they fit
        num_fat_sectors = num_difat_sectors = 0
        while True:
            total = num_data_sectors + num_fat_sectors + num_difat_sectors
            fat_needed = -(-total // ENTRIES_PER_SECTOR)
            difat_needed = -(
                -max(0, fat_needed - NUM_HEADER_DIFAT) // (ENTRIES_PER_SECTOR - 1)
            )

            if (fat_needed, difat_needed) == (num_fat_sectors, num_difat_sectors):
                break

            num_fat_sectors, num_difat_sectors = fat_needed, difat_needed

        first_fat_sector = num_data_sectors
        first_difat_sector = first_fat_sector + num_fat_sectors
        fat_sectors = list(range(first_fat_sector, first_difat_sector))

        # Write FAT, marking FAT and DIFAT sectors
        self.fat.extend([FATSECT] * num_fat_sectors)
        self.fat.extend([DIFSECT] * num_difat_sectors)
        self.fp.write(_pack_sectors(self.fat))

        # Write DIFAT sectors for FAT sectors that do not fit in the header
        overflow = fat_sectors[NUM_HEADER_DIFAT:]
        per_sector = ENTRIES_PER_SECTOR - 1

        for index in range(num_difat_sectors):
            entries = overflow[index * per_sector : (index + 1) * per_sector]
            entries.extend([FREESECT] * (per_sector - len(entries)))

            if index + 1 < num_difat_sectors:
                entries.append(first_difat_sector + index + 1)
            else:
                entries.append(ENDOFCHAIN)

            self.fp.write(_pack_sectors(entries))

        header_difat = fat_sectors[:NUM_HEADER_DIFAT]
        header_difat.extend([FREESECT] * (NUM_HEADER_DIFAT - len(header_difat)))

        # Write header
        self.fp.seek(0)
        self.fp.write(
            HEADER_FORMAT.pack(
                HEADER_SIGNATURE,
                bytes(16),  # Header CLSID
                MINOR_VERSION,
                MAJOR_VERSION,
                BYTE_ORDER,
                SECTOR_SHIFT,
                MINI_SECTOR_SHIFT,
                0,  # Number of directory sectors, must be zero for major version 3
                num_fat_sectors,
                header["first_dir_sector"],
                0,  # Transaction signature number
                MINI_STREAM_CUTOFF_SIZE,
                header["first_mini_fat_sector"],
                header["num_mini_fat_sectors"],
                first_difat_sector if num_difat_sectors else ENDOFCHAIN,
                num_difat_sectors,
                *header_difat,
            )
        )


def _walk(src, dest):
    """
    Yield `(parent, name, path)` for every directory and file under `src`.

    `parent` is the relative path of the parent directory (empty for `src` itself) and `path` is `None` for directories. Entries are sorted so output is deterministic. `dest` is skipped if it lies inside `src`.

    """

    dest = os.path.abspath(dest)

    for dirpath, dirnames, filenames in os.walk(src):
        dirnames.sort()
        filenames.sort()

        rel = os.path.relpath(dirpath, src)
        parent = "" if rel == os.curdir else rel

        if parent:
            yield os.path.dirname(parent), os.path.basename(parent), None

        for name in filenames:
            path = os.path.join(dirpath, name)

            if os.path.abspath(path) != dest:
                yield parent, name, path


def compress(src, dest=None):
    """
    Create compound file from given directory tree.

    """

    src = os.path.normpath(src)

    if not os.path.isdir(src):
        raise NotADirectoryError(f"Source is not a directory: {src}")

    if not dest:
        dest = f"{src}.cfb"

    try:
        _compress(src, dest)

    except BaseException:
        # Do not leave a partially written file behind
        if os.path.exists(dest):
            os.remove(dest)

        raise


def _compress(src, dest):
    root = directory.Entry("Root Entry", obj_type=directory.OBJTY_ROOT)
    storages = {"": root}
    mini_stream = _MiniStream()

    with open(dest, "wb") as fp:
        writer = _SectorWriter(fp)

        for parent, name, path in _walk(src, dest):
            storage = storages[parent]

            # Add storage as directory entry
            if path is None:
                storages[os.path.join(parent, name)] = storage.add(
                    directory.Entry(name, obj_type=directory.OBJTY_STORAGE),
                )
                continue

            entry = storage.add(directory.Entry(name, obj_type=directory.OBJTY_STREAM))

            # Write data into mini-sector if smaller than cutoff, otherwise into sector
            with open(path, "rb") as f:
                head = f.read(MINI_STREAM_CUTOFF_SIZE)

                if len(head) < MINI_STREAM_CUTOFF_SIZE:
                    entry.sector, entry.size = mini_stream.write(head), len(head)

                else:
                    entry.sector, entry.size = writer.copy(head, f)

        # Write mini stream into sector; its location is stored in root entry
        root.sector = writer.write(mini_stream.data)
        root.size = len(mini_stream.data)

        # Write mini-fat into sector
        mini_fat_data = _pack_sectors(mini_stream.fat) if mini_stream.fat else b""
        first_mini_fat_sector = writer.write(mini_fat_data)

        # Write directory entry into sector
        first_dir_sector = writer.write(directory.serialize(root, SECTOR_SIZE))

        # Write fat, difat and header
        writer.close(
            first_dir_sector=first_dir_sector,
            first_mini_fat_sector=first_mini_fat_sector,
            num_mini_fat_sectors=len(mini_fat_data) // SECTOR_SIZE,
        )


def _safe_name(name):
    """
    Return entry name as a file name, rejecting names that would escape the destination directory.

    """

    if (
        name in ("", os.curdir, os.pardir)
        or os.sep in name
        or (os.altsep and os.altsep in name)
        or "\0" in name
        or os.path.isabs(name)
        or os.path.splitdrive(name)[0]
    ):
        raise ValueError(f"Unsafe entry name in compound file: {name!r}")

    return name


def decompress(src, dest=None):
    """
    Decompress compound file and create directory at destination.

    Raises `CompoundFileError` if `src` is not a valid compound file.

    """

    # Set path same as `src` if dest is not set
    if not dest:
        dest, _ = os.path.splitext(src)

    # Open before creating `dest`, so an invalid file leaves nothing behind
    with reader.CompoundFileReader(src) as cfb:
        os.makedirs(dest, exist_ok=True)

        # Storages are yielded before their children
        for parts, entry in cfb.walk():
            path = os.path.join(dest, *parts[:-1], _safe_name(parts[-1]))

            if entry.obj_type == directory.OBJTY_STORAGE:
                os.makedirs(path, exist_ok=True)

            else:
                with open(path, "wb") as f:
                    cfb.copy_stream(entry, f)

    # Return True if decompressed successfully
    return True


class CompoundFile:
    compress = staticmethod(compress)
    decompress = staticmethod(decompress)

    def open(self, src):
        """
        Open compound file and create CompoundFile instance.

        """

        raise NotImplementedError
