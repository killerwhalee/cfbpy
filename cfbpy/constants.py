import array, struct

# Header constants
HEADER_SIGNATURE = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
MINOR_VERSION = 0x003E
MAJOR_VERSION = 0x0003
BYTE_ORDER = 0xFFFE
SECTOR_SHIFT = 0x0009  # 512-byte sectors for major version 3
MINI_SECTOR_SHIFT = 0x0006
MINI_STREAM_CUTOFF_SIZE = 0x00001000

HEADER_FORMAT = struct.Struct("<8s16sHHHHH6xIIIIIIIII109I")
HEADER_SIZE = HEADER_FORMAT.size  # 512 bytes
NUM_HEADER_DIFAT = 109

# Sector constants
MAXREGSECT = 0xFFFFFFFA  #: (-6) maximum SECT
DIFSECT = 0xFFFFFFFC  #: (-4) denotes a DIFAT sector in a FAT
FATSECT = 0xFFFFFFFD  #: (-3) denotes a FAT sector in a FAT
ENDOFCHAIN = 0xFFFFFFFE  #: (-2) end of a virtual stream chain
FREESECT = 0xFFFFFFFF  #: (-1) unallocated sector

# Array typecode of unsigned 32-bit integer, used for FAT and MINIFAT
UINT32 = next(code for code in "IL" if array.array(code).itemsize == 4)
