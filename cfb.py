from olefile import OleFileIO

import os, struct

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
STGTY_EMPTY = 0  #: empty directory entry
STGTY_STORAGE = 1  #: element is a storage object
STGTY_STREAM = 2  #: element is a stream object
STGTY_LOCKBYTES = 3  #: element is an ILockBytes object
STGTY_PROPERTY = 4  #: element is an IPropertyStorage object
STGTY_ROOT = 5  #: element is a root storage


class CompoundFile:
    class Directory:
        def __init__(self) -> None:
            # Directory data
            self.name = b""
            self.type = STGTY_EMPTY
            self.subdirs = []
            self.sector = FREESECT
            self.size = 0

    def __init__(self) -> None:
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
        self.sector = b""
        self.mini_sector = b""

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
            self.sector += head
            self.fat.append(len(self.fat) + 1)

            head, tail = tail[: self.sector_shift], tail[self.sector_shift :]

        # Write last head to sector and end FAT chain
        self.sector += struct.pack("<self.sector_shifts", head)
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
        head, tail = data[:64], data[64:]

        # Write head to sector right before last head
        while tail:
            self.mini_sector += head
            self.mini_fat.append(len(self.mini_fat) + 1)

            head, tail = tail[:64], tail[64:]

        # Write last head to sector and end MINIFAT chain
        self.mini_sector += struct.pack("<64s", head)
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
            self.difat.append(len(self.sector) // self.sector_shift)
            self.sector += struct.pack("<self.sector_shifts", head)

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

        # Open file path to save data
        fp = open(dest, "wb")

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
        fp.write(header_data)

        # Write difat
        fp.write(b"".join([struct.pack("<I", entry) for entry in self.difat]))

        # Write sectors
        fp.write(self.sector)

        # Save file
        fp.close()

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
            storage_name = os.path.relpath(root, src).lstrip(".")

            # Add storage as directory entry
            if storage_name:
                raise NotImplementedError

            else:
                raise NotImplementedError

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

                # Add stream as directory entry
                raise NotImplementedError

        # Write mini-sector into sector
        mini_sector_index = cfb.write_sector(cfb.mini_sector)

        # Write mini-fat into sector
        minifat_data = b"".join(
            [struct.pack("<I", entry) for entry in cfb.mini_fat],
        )
        cfb.num_mini_fat_sectors = (
            len(minifat_data) // cfb.sector_shift
            + bool(len(minifat_data) % cfb.sector_shift),
        )
        cfb.first_mini_fat_sector = cfb.write_sector(minifat_data)

        # Write directory entry into sector
        directory_data = b""
        cfb.num_dir_sectors = (
            len(directory_data) // cfb.sector_shift
            + bool(len(directory_data) % cfb.sector_shift),
        )
        cfb.first_dir_sector = cfb.write_sector(directory_data)

        # Write fat into sector
        fat_data = b"".join([struct.pack("<I", entry) for entry in cfb.fat])
        cfb.write_fat(fat_data)
        cfb.num_fat_sectors = (
            len(fat_data) // cfb.sector_shift + bool(len(fat_data) % cfb.sector_shift),
        )

        # Export results as file
        cfb.save(dest)


if __name__ == "__main__":
    CompoundFile.compress("tests/test")
