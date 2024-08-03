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

        # Sector
        self.sector = b""
        self.mini_sector = b""

    def write_stream(self, data: bytes):
        """
        Write data stream to the sector.

        """

        # Get sector index for data
        sector_index = len(self.fat)

        # Add zero-padding for data
        padding = bytes((512 - len(data) % 512) % 512)
        tail = data + padding

        # Separate tail into head/tail
        head, tail = tail[:512], tail[512:]

        # Write head to sector right before last head
        while tail:
            self.sector += head
            self.fat.append(len(self.fat) + 1)

            head, tail = tail[:512], tail[512:]

        # Write last head to sector and end FAT chain
        self.sector += head
        self.fat.append(ENDOFCHAIN)

        # Return first sector number of stream
        return sector_index

    def write_mini_stream(self, data: bytes):
        """
        Write mini data stream to the sector.

        """

        # Get sector index for data
        sector_index = len(self.mini_fat)

        # Add zero-padding for data
        padding = bytes((64 - len(data) % 64) % 64)
        tail = data + padding

        # Separate tail into head/tail
        head, tail = tail[:64], tail[64:]

        # Write head to sector right before last head
        while tail:
            self.mini_sector += head
            self.mini_fat.append(len(self.mini_fat) + 1)

            head, tail = tail[:64], tail[64:]

        # Write last head to sector and end FAT chain
        self.mini_sector += head
        self.mini_fat.append(ENDOFCHAIN)

        # Return first sector number of stream
        return sector_index

    def open(self, dest):
        """
        Open compound file and create CompoundFile instance.

        """

        raise NotImplementedError

    def save(self, dest):
        """
        Save compound file from CompoundFile instance.

        """

        # Open file path to save data
        f = open(dest, "wb")

        # Write header sector
        data = struct.pack_into(
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
        f.write(data)

        for difat in self.difat:
            data = struct.pack("<I", difat)
            f.write(data)

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
                raise NotImplementedError

            else:
                pass

            for stream in streams:
                # Read data from file
                with open(f"{root}/{stream}", "rb") as f:
                    data = f.read()
                    size = len(data)

                # Write data as stream
                if size > 4096:
                    cfb.write_stream(data)

                else:
                    cfb.write_mini_stream(data)

                # Add stream as directory entry
                raise NotImplementedError


if __name__ == "__main__":
    CompoundFile.compress("tests/test")
