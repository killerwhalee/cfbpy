from olefile import OleFileIO

import os, struct
import math

# Header constants
HEADER_SIGNATURE = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
MINOR_VERSION = 0x003E
MAJOR_VERSION = 0x0003
BYTE_ORDER = 0xFFEE
MINI_SECTOR_SHIFT = 0x0006
MINI_STREAM_CUTOFF_SIZE = 0x00001000

# Sector constants
MAXREGSECT = 0xFFFFFFFA  #: (-6) maximum SECT
DIFSECT = 0xFFFFFFFC  #: (-4) denotes a DIFAT sector in a FAT
FATSECT = 0xFFFFFFFD  #: (-3) denotes a FAT sector in a FAT
ENDOFCHAIN = 0xFFFFFFFE  #: (-2) end of a virtual stream chain
FREESECT = 0xFFFFFFFF  #: (-1) unallocated sector

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


class CompoundFile:
    class Directory:
        def __init__(self) -> None:
            # Directory data
            self.name = "".encode("utf-16-le")
            self.type = OBJTY_EMPTY
            self.subdirs = []
            self.sector = 0x00000000
            self.size = 0

        def __bytes__(self) -> bytes:
            directory_data = struct.pack(
                "<64sHBBIII16sIQQIQ",
                self.name,
                len(self.name) + 1,
                self.type,
                0x01,
                0xFFFFFFFF,
                0xFFFFFFFF,
                0xFFFFFFFF,
                bytes(16),
                0x00000000,
                0x0000000000000000,
                0x0000000000000000,
                self.sector,
                self.size,
            )

            return directory_data

        def find(self, path):
            head, *tail = path.split("/", 1)

            target_dir = None

            for subdir in self.subdirs:
                if subdir.name == head:
                    target_dir = subdir
                    break

            # Return target directory
            # if no target was found or no tail is left
            if not target_dir or not tail:
                return target_dir

            # Recursively find target directory
            return target_dir.find(tail[0])

    def __init__(self) -> None:
        # File metadata
        self.file_name = ""

        # Header data
        self.header_signature = HEADER_SIGNATURE
        self.header_clsid = 0
        self.minor_version = MINOR_VERSION
        self.major_version = MAJOR_VERSION
        self.byte_order = BYTE_ORDER
        self.sector_shift = 0x0009 if self.major_version == 0x0003 else 0x000C
        self.mini_sector_shift = MINI_SECTOR_SHIFT
        self.num_dir_sectors = 0
        self.num_fat_sectors = 0
        self.first_dir_sector = 0
        self.transaction_signature_number = 0
        self.mini_stream_cutoff_size = MINI_STREAM_CUTOFF_SIZE
        self.first_mini_fat_sector = 0
        self.num_mini_fat_sectors = 0
        self.first_difat_sector = 0
        self.num_difat_sectors = 0

        # File Allocation Table (FAT)
        self.fat = []
        self.mini_fat = []
        self.difat = []

        # Sector data
        self.sectors = []
        self.mini_sectors = []

        # Directories
        self.root_directory = CompoundFile.Directory()
        self.root_directory.name = "Root Entry".encode("utf-16-le")
        self.root_directory.type = OBJTY_ROOT

    def write_sector(self, data: bytes):
        """
        Write data stream to the sector.

        """

        # Get sector index for data
        sector_index = len(self.fat)

        # Separate data into head/tail
        head, tail = data[: self.sector_shift], data[self.sector_shift :]

        # Write head to sector right before last head
        while tail:
            self.sectors.append(head)
            self.fat.append(len(self.sectors))

            head, tail = tail[: self.sector_shift], tail[self.sector_shift :]

        # Write last head to sector and end FAT chain
        self.sectors.append(struct.pack(f"<{self.sector_shift}s", head))
        self.fat.append(ENDOFCHAIN)

        # Return first sector number of stream
        return sector_index

    def write_mini_sector(self, data: bytes):
        """
        Write mini data stream to the sector.

        """

        # Get sector index for data
        sector_index = len(self.mini_fat)

        # Separate data into head/tail
        head, tail = data[: self.mini_sector_shift], data[self.mini_sector_shift :]

        # Write head to sector right before last head
        while tail:
            self.mini_sectors.append(head)
            self.mini_fat.append(len(self.mini_sectors))

            head, tail = (
                tail[: self.mini_sector_shift],
                tail[self.mini_sector_shift :],
            )

        # Write last head to sector and end MINIFAT chain
        self.mini_sectors.append(
            struct.pack(f"<{self.mini_sector_shift}s", head),
        )
        self.mini_fat.append(ENDOFCHAIN)

        # Return first sector number of stream
        return sector_index

    def write_fat(self, data: bytes):
        """
        Write fat to the sector and write difat array.

        """

        # Separate data into head/tail
        head, tail = data[: self.sector_shift], data[self.sector_shift :]

        # Write head to sector until there is no head left
        while head:
            self.difat.append(len(self.sectors))
            self.sectors.append(struct.pack(f"<{self.sector_shifts}", head))

            head, tail = tail[: self.sector_shift], tail[self.sector_shift :]

    def open(self, src):
        """
        Open compound file and create CompoundFile instance.

        """

        raise NotImplementedError

    def save(self, dest):
        """
        Save compound file from CompoundFile instance.

        """
        # Update header data
        self.num_fat_sectors = None
        self.num_mini_fat_sectors = None
        self.num_difat_sectors = None
        self.num_dir_sectors = None

        # Write header sector
        header_data = struct.pack_into(
            "<8s16xHHHHH6xIIIIIIIII",
            self.header_signature,
            self.minor_version,
            self.major_version,
            self.byte_order,
            self.sector_shift,
            self.mini_sector_shift,
            self.num_dir_sectors,
            self.num_fat_sectors,
            self.first_dir_sector,
            self.transaction_signature_number,
            self.mini_stream_cutoff_size,
            self.first_mini_fat_sector,
            self.num_mini_fat_sectors,
            self.first_difat_sector,
            self.num_difat_sectors,
        )

        # Open file path to save data
        with open(dest, "wb") as fp:
            # Write header
            fp.write(header_data)

            # Write difat
            fp.write(b"".join([struct.pack("<I", entry) for entry in self.difat]))

            # Write sectors
            fp.write(b"".join(self.sectors))

    @staticmethod
    def decompress(src, dest=None):
        """
        Decompress compound file and create directory at destination.

        """

        # Set path same as `src` if dest is not set
        if not dest:
            dest, _ = os.path.splitext(src)

        os.makedirs(dest, exist_ok=True)

        # Open compound file
        olefile = OleFileIO(src)

        # Iterate every stream in compound file
        for *storage_path, stream_path in olefile.listdir():
            # Create storge as directory
            full_path = dest + "/" + "/".join(storage_path)
            os.makedirs(full_path, exist_ok=True)

            # Read stream and export as file
            real_path = full_path + "/" + stream_path

            with open(real_path, "wb") as f:
                data_path = "/".join(storage_path + [stream_path])
                data = olefile.openstream(data_path).read()

                f.write(data)

        # Return True if decompressed successfully
        return True

    @staticmethod
    def compress(src, dest=None):
        """
        Create compound file from given directory tree.

        """

        cfb = CompoundFile()

        if not dest:
            dest = f"{src}.cfb"

        for root, _, streams in os.walk(src):
            storage = os.path.relpath(root, src).lstrip(".")

            # Add storage as directory entry
            if storage:
                *path, storage_name = storage.split("/")

                # Initialize directory entry for storage
                storage_dir = CompoundFile.Directory()
                storage_dir.name = storage_name.encode("utf-16-le")
                storage_dir.type = OBJTY_STORAGE

                # Add storage as subdir of parent storage
                parent_dir = cfb.root_directory.find("/".join(path))
                parent_dir.subdirs.append(storage_dir)

            # Use root storage if storage name is empty
            else:
                storage_dir = cfb.root_directory

            for stream in streams:
                # Read data from file
                with open(f"{root}/{stream}", "rb") as f:
                    stream_data = f.read()
                    stream_size = len(stream_data)

                # Write data into sector or mini-sector
                if stream_size > MINI_STREAM_CUTOFF_SIZE:
                    stream_sector_index = cfb.write_sector(stream_data)

                else:
                    stream_sector_index = cfb.write_mini_sector(stream_data)

                # Initialize directory entry for stream
                stream_dir = CompoundFile.Directory()
                stream_dir.name = stream.encode("utf-16-le")
                stream_dir.type = OBJTY_STREAM
                stream_dir.size = stream_size
                stream_dir.sector = stream_sector_index

                # Add stream as subdir of storage
                storage_dir.subdirs.append(stream_dir)

        # Write mini-sector into sector
        mini_sector_data = b"".join(cfb.mini_sectors)
        mini_sector_index = cfb.write_sector(mini_sector_data)

        cfb.root_directory.size = len(mini_sector_data)
        cfb.root_directory.sector = mini_sector_index

        # Write mini-fat into sector
        mini_fat_data = b"".join(
            [struct.pack("<I", entry) for entry in cfb.mini_fat],
        )
        cfb.num_mini_fat_sectors = math.ceil(
            len(mini_fat_data) / cfb.sector_shift,
        )
        cfb.first_mini_fat_sector = cfb.write_sector(mini_fat_data)

        # Write directory entry into sector
        directory_data = b""
        cfb.num_dir_sectors = math.ceil(len(directory_data) / cfb.sector_shift)
        cfb.first_dir_sector = cfb.write_sector(directory_data)

        # Write fat into sector
        fat_data = b"".join([struct.pack("<I", entry) for entry in cfb.fat])
        cfb.write_fat(fat_data)
        cfb.num_fat_sectors = math.ceil(len(fat_data) / cfb.sector_shift)

        # Export results as file
        cfb.save(dest)


if __name__ == "__main__":
    CompoundFile.compress("tests/test")
