from olefile import OleFileIO
import os, struct
import math

import directory

# Header constants
HEADER_SIGNATURE = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
MINOR_VERSION = 0x003E
MAJOR_VERSION = 0x0003
BYTE_ORDER = 0xFFFE
MINI_SECTOR_SHIFT = 0x0006
MINI_STREAM_CUTOFF_SIZE = 0x00001000

# Sector constants
MAXREGSECT = 0xFFFFFFFA  #: (-6) maximum SECT
DIFSECT = 0xFFFFFFFC  #: (-4) denotes a DIFAT sector in a FAT
FATSECT = 0xFFFFFFFD  #: (-3) denotes a FAT sector in a FAT
ENDOFCHAIN = 0xFFFFFFFE  #: (-2) end of a virtual stream chain
FREESECT = 0xFFFFFFFF  #: (-1) unallocated sector


class CompoundFile:
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
        self.root_directory = directory.Directory()
        self.num_dir = 1  # Starts from 1 due to root entry

        # Insert root entry to root directory
        self.root_directory.insert(
            directory.EntryData(
                "Root Entry",
                stream_id=0,
                obj_type=directory.OBJTY_ROOT,
            ),
        )

    def search_directory(self, path):
        """
        Search path in directory

        Returns entry with given path if found. Return `None` if there is no such path in directory.

        """

        current_entry = self.root_directory.root

        for name in path.split("/"):
            # Skip if name is blank (double slash or root directory)
            if name == "":
                continue

            # Search name for current entry
            current_entry = current_entry.data.child.search_name(name=name)

            # Return `None` if current_entry is NIL
            if current_entry.data is None:
                return None

        # Return search result
        return current_entry

    def insert_directory(self, path, data: directory.EntryData):
        """
        Insert new entry as child of other entry with given path.

        """

        if (entry := self.search_directory(path)) is None:
            raise Exception(f"Given path not found: {path}")

        # Update entry data
        data.stream_id = self.num_dir
        data.parent = entry

        # Insert new entry as child of parent entry
        entry.data.child.insert(data)
        self.num_dir += 1

    def write_sector(self, data: bytes):
        """
        Write data stream to the sector.

        """

        # Get sector index for data
        sector_index = len(self.fat)

        # Get sector size
        sector_size = 1 << self.sector_shift

        # Separate data into head/tail
        head, tail = data[:sector_size], data[sector_size:]

        # Write head to sector right before last head
        while tail:
            self.sectors.append(head)
            self.fat.append(len(self.sectors))

            head, tail = tail[:sector_size], tail[sector_size:]

        # Write last head to sector and end FAT chain
        self.sectors.append(struct.pack(f"<{sector_size}s", head))
        self.fat.append(ENDOFCHAIN)

        # Return first sector number of stream
        return sector_index

    def write_mini_sector(self, data: bytes):
        """
        Write mini data stream to the sector.

        """

        # Get sector index for data
        sector_index = len(self.mini_fat)

        # Get mini sector size
        mini_sector_size = 1 << self.mini_sector_shift

        # Separate data into head/tail
        head, tail = data[:mini_sector_size], data[mini_sector_size:]

        # Write head to sector right before last head
        while tail:
            self.mini_sectors.append(head)
            self.mini_fat.append(len(self.mini_sectors))

            head, tail = tail[:mini_sector_size], tail[mini_sector_size:]

        # Write last head to sector and end MINIFAT chain
        self.mini_sectors.append(
            struct.pack(f"<{mini_sector_size}s", head),
        )
        self.mini_fat.append(ENDOFCHAIN)

        # Return first sector number of stream
        return sector_index

    def write_fat(self, data: bytes):
        """
        Write fat to the sector and write difat array.

        """

        # Get sector size
        sector_size = 1 << self.sector_shift

        # Separate data into head/tail
        head, tail = data[:sector_size], data[sector_size:]

        # Write head to sector until there is no head left
        while head:
            self.difat.append(len(self.sectors))
            self.sectors.append(struct.pack(f"<{sector_size}s", head))

            head, tail = tail[:sector_size], tail[sector_size:]

    def open(self, src):
        """
        Open compound file and create CompoundFile instance.

        """

        raise NotImplementedError

    def save(self, dest):
        """
        Save compound file from CompoundFile instance.

        """

        # Write header sector
        header_data = struct.pack(
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
            for entry in self.difat:
                fp.write(struct.pack("<I", entry))

            # Patch rest of difat entry
            for _ in range(109 - len(self.difat)):
                fp.write(struct.pack("<I", 0xFFFFFFFF))

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

            # Add storage as directory entry is storage is not a root
            if storage:
                *path, storage_name = storage.split("/")

                # Insert storage into directory
                cfb.insert_directory(
                    "/".join(path),
                    directory.EntryData(
                        name=storage_name,
                        obj_type=directory.OBJTY_STORAGE,
                    ),
                )

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

                # Insert stream into directory
                cfb.insert_directory(
                    storage,
                    directory.EntryData(
                        name=stream,
                        obj_type=directory.OBJTY_STREAM,
                        sector=stream_sector_index,
                        size=stream_size,
                    ),
                )

        # Write mini-sector into sector
        mini_sector_data = b"".join(cfb.mini_sectors)
        mini_sector_index = cfb.write_sector(mini_sector_data)

        cfb.root_directory.root.data.size = len(mini_sector_data)
        cfb.root_directory.root.data.sector = mini_sector_index

        # Get sector size
        sector_size = 1 << cfb.sector_shift

        # Write mini-fat into sector
        mini_fat_data = b"".join(
            [struct.pack("<I", entry) for entry in cfb.mini_fat],
        )
        cfb.num_mini_fat_sectors = math.ceil(
            len(mini_fat_data) / sector_size,
        )
        cfb.first_mini_fat_sector = cfb.write_sector(mini_fat_data)

        # Round up to nearest multiple of 4
        directory_list = [b"\xff\xff\xff\xff"] * ((cfb.num_dir + 3) // 4) * 4

        # Insert directory entry into directory list
        for entry in cfb.root_directory.traverse():
            directory_list[entry.stream_id()] = bytes(entry)

        # Write directory entry into sector
        directory_data = b"".join(directory_list)
        cfb.num_dir_sectors = math.ceil(len(directory_data) / sector_size)
        cfb.first_dir_sector = cfb.write_sector(directory_data)

        # Write fat into sector
        fat_data = b"".join([struct.pack("<I", entry) for entry in cfb.fat])
        cfb.write_fat(fat_data)
        cfb.num_fat_sectors = math.ceil(len(fat_data) / sector_size)

        # Hotfixed instance variables
        # TODO: You have to remove this line after problem is fully recognized.
        cfb.first_difat_sector = 0xFFFFFFFE
        cfb.num_dir_sectors = 0

        # Export results as file
        cfb.save(dest)
