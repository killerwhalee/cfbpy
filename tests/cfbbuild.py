"""
Test-only compound file writer that produces layouts cfbpy itself never writes.

- Major version 3 (512-byte sectors) or 4 (4096-byte sectors)
- Sectors and mini sectors allocated in random order, so every chain is fragmented
- Directory stream IDs shuffled, and siblings linked as a degenerate tree (each entry's right sibling is the next name), which is valid but maximally deep
- Optionally, garbage in the high 32 bits of stream sizes, which version 3 readers must ignore

It shares no code with cfbpy, so reader bugs cannot be masked by the same bug in the writer.

"""

import random
import struct

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


def sort_key(name):
    upper = "".join(c.upper() if len(c.upper()) == 1 else c for c in name)
    return len(name.encode("utf-16-le")), upper.encode("utf-16-be")


def build(path, files, version=3, seed=0, size_high=0):
    """
    Write a compound file to `path`.

    `files` maps "a/b/c" paths to stream data, or to `None` for an (empty) storage.

    """

    rng = random.Random(seed)
    sector_size = 512 if version == 3 else 4096

    # Directory tree, including implicit parent storages
    nodes = {(): {"name": "Root Entry", "type": 5, "children": [], "data": None}}
    for file_path, data in files.items():
        parts = tuple(file_path.split("/"))
        for depth in range(1, len(parts) + 1):
            key = parts[:depth]
            if key not in nodes:
                is_leaf = depth == len(parts)
                nodes[key] = {
                    "name": parts[depth - 1],
                    "type": 2 if is_leaf and data is not None else 1,
                    "children": [],
                    "data": data if is_leaf else None,
                }
                nodes[parts[: depth - 1]]["children"].append(key)

    # Shuffled stream IDs, root stays 0
    keys = [key for key in nodes if key]
    rng.shuffle(keys)
    sid = {(): 0}
    sid.update({key: index + 1 for index, key in enumerate(keys)})

    # Mini stream with shuffled mini sector allocation
    small = [key for key in keys if nodes[key]["type"] == 2 and 0 < len(nodes[key]["data"]) < 4096]
    mini_sectors = []
    for key in small:
        data = nodes[key]["data"]
        mini_sectors += [(key, i) for i in range(-(-len(data) // 64))]
    order = list(range(len(mini_sectors)))
    rng.shuffle(order)
    mini_stream = bytearray(len(mini_sectors) * 64)
    mini_fat = [FREESECT] * len(mini_sectors)
    mini_start = {}
    chains = {}
    for slot, (key, index) in zip(order, mini_sectors):
        data = nodes[key]["data"]
        mini_stream[slot * 64 : slot * 64 + 64] = data[index * 64 : index * 64 + 64].ljust(64, b"\0")
        chains.setdefault(key, []).append(slot)
    for key, slots in chains.items():
        mini_start[key] = slots[0]
        for current, following in zip(slots, slots[1:] + [ENDOFCHAIN]):
            mini_fat[current] = following

    def pad(data):
        return bytes(data) + bytes(-len(data) % sector_size)

    mini_fat += [FREESECT] * (-len(mini_fat) % (sector_size // 4))
    mini_fat_data = struct.pack(f"<{len(mini_fat)}I", *mini_fat)

    # Directory entries; siblings form a right-leaning chain in sorted order
    for key, node in nodes.items():
        children = sorted(node["children"], key=lambda child: sort_key(nodes[child]["name"]))
        node["child"] = sid[children[0]] if children else NOSTREAM
        for current, following in zip(children, children[1:] + [None]):
            nodes[current]["right"] = sid[following] if following else NOSTREAM
    nodes[()]["right"] = NOSTREAM

    # Regular-sector blocks: (label, data)
    blocks = [("mini_stream", bytes(mini_stream)), ("mini_fat", mini_fat_data)]
    blocks += [(key, nodes[key]["data"]) for key in keys if nodes[key]["type"] == 2 and len(nodes[key]["data"]) >= 4096]
    blocks.append(("directory", None))  # filled after start sectors are known

    num_dir_sectors = -(-len(nodes) * 128 // sector_size)
    counts = [-(-len(data) // sector_size) if data is not None else num_dir_sectors for _, data in blocks]
    num_data = sum(counts)

    num_fat = 0
    while num_fat * (sector_size // 4) < num_data + num_fat:
        num_fat += 1
    assert num_fat <= 109, "test builder does not write DIFAT sectors"

    total = num_data + num_fat
    slots = list(range(total))
    rng.shuffle(slots)
    fat_sectors, slots = slots[:num_fat], slots[num_fat:]

    fat = [FREESECT] * (num_fat * sector_size // 4)
    for sector in fat_sectors:
        fat[sector] = FATSECT

    sectors = {}
    start = {}
    position = 0
    for (label, _), count in zip(blocks, counts):
        assigned = slots[position : position + count]
        position += count
        start[label] = assigned[0] if assigned else ENDOFCHAIN
        for current, following in zip(assigned, assigned[1:] + [ENDOFCHAIN]):
            fat[current] = following
        sectors[label] = assigned

    def entry_bytes(key, node):
        if key == ():
            sector, size = start["mini_stream"], len(mini_stream)
        elif node["type"] == 1:
            sector, size = 0, 0
        elif not node["data"]:
            sector, size = ENDOFCHAIN, 0
        elif len(node["data"]) < 4096:
            sector, size = mini_start[key], len(node["data"])
        else:
            sector, size = start[key], len(node["data"])
        if size and version == 3:
            size |= size_high << 32
        name = node["name"].encode("utf-16-le")
        return struct.pack(
            "<64sHBBIII16sIQQIQ",
            name, len(name) + 2, node["type"], 1,  # all black
            NOSTREAM, node["right"], node["child"],
            bytes(16), 0, 0, 0, sector, size,
        )

    empty = bytes(64) + struct.pack("<HBBIII", 0, 0, 0, NOSTREAM, NOSTREAM, NOSTREAM) + bytes(48)
    directory = b"".join(entry_bytes(key, nodes[key]) for key in sorted(nodes, key=lambda k: sid[k]))
    directory += empty * (num_dir_sectors * sector_size // 128 - len(nodes))

    data_by_label = dict(blocks)
    data_by_label["directory"] = directory

    image = bytearray(total * sector_size)
    for label, assigned in sectors.items():
        data = pad(data_by_label[label])
        for index, sector in enumerate(assigned):
            image[sector * sector_size : (sector + 1) * sector_size] = data[index * sector_size : (index + 1) * sector_size]
    fat_data = struct.pack(f"<{len(fat)}I", *fat)
    for index, sector in enumerate(fat_sectors):
        image[sector * sector_size : (sector + 1) * sector_size] = fat_data[index * sector_size : (index + 1) * sector_size]

    header_difat = fat_sectors + [FREESECT] * (109 - num_fat)
    header = struct.pack(
        "<8s16sHHHHH6xIIIIIIIII109I",
        b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1",
        bytes(16),
        0x3E,
        version,
        0xFFFE,
        9 if version == 3 else 12,
        6,
        0 if version == 3 else num_dir_sectors,
        num_fat,
        start["directory"],
        0,
        4096,
        start["mini_fat"],
        len(sectors["mini_fat"]),
        ENDOFCHAIN,
        0,
        *header_difat,
    )

    with open(path, "wb") as f:
        f.write(header.ljust(sector_size, b"\0"))
        f.write(image)
