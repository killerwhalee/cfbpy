import os
import struct
import subprocess
import sys

import olefile
import pytest

from cfbpy import CompoundFile, CompoundFileError

import cfbbuild
from test_cfb import read_tree, expected_entries

FILES = {
    "tiny": b"x",
    "empty": b"",
    "mini/one": os.urandom(64),
    "mini/edge": os.urandom(4095),
    "mini/odd": os.urandom(1000),
    "big/exact": os.urandom(4096),
    "big/fragmented": os.urandom(50_000),
    "big/deep/nested/stream": os.urandom(9000),
    "empty_storage": None,
}


def expected(files):
    return expected_entries(
        {path: data for path, data in files.items() if data is not None},
        [path for path, data in files.items() if data is None],
    )


def decompress_tree(tmp_path, src):
    out = tmp_path / "out"
    assert CompoundFile.decompress(str(src), str(out))
    return read_tree(out)


@pytest.mark.parametrize("version", [3, 4])
@pytest.mark.parametrize("seed", range(5))
def test_fragmented_file(tmp_path, version, seed):
    src = tmp_path / "in.cfb"
    cfbbuild.build(src, FILES, version=version, seed=seed)

    # Make sure the test builder itself writes a valid file
    with olefile.OleFileIO(str(src), raise_defects=olefile.DEFECT_INCORRECT) as ole:
        for path, data in FILES.items():
            if data is not None:
                assert ole.openstream(path).read() == data

    assert decompress_tree(tmp_path, src) == expected(FILES)


def test_ignores_size_high_bits_in_version_3(tmp_path):
    src = tmp_path / "in.cfb"
    cfbbuild.build(src, FILES, version=3, size_high=0xDEADBEEF)

    assert decompress_tree(tmp_path, src) == expected(FILES)


def test_degenerate_tree(tmp_path):
    # Thousands of siblings linked as a single chain must not hit recursion limits
    files = {f"s/n{i:05d}": bytes([i % 256]) for i in range(5000)}
    src = tmp_path / "in.cfb"
    cfbbuild.build(src, files)

    assert decompress_tree(tmp_path, src) == expected(files)


def test_olefile_not_required():
    code = "import sys, cfbpy; assert 'olefile' not in sys.modules"
    subprocess.run([sys.executable, "-c", code], check=True)


@pytest.mark.parametrize(
    "data",
    [b"", b"\xD0\xCF\x11\xE0", b"not a compound file" * 100],
    ids=["empty", "short", "garbage"],
)
def test_not_a_compound_file(tmp_path, data):
    src = tmp_path / "in.cfb"
    src.write_bytes(data)

    with pytest.raises(CompoundFileError):
        CompoundFile.decompress(str(src), str(tmp_path / "out"))

    assert not (tmp_path / "out").exists()


def make_compound_file(tmp_path, files):
    src = tmp_path / "src"
    for path, data in files.items():
        full_path = src.joinpath(*path.split("/"))
        full_path.parent.mkdir(parents=True, exist_ok=True)
        full_path.write_bytes(data)

    dest = tmp_path / "in.cfb"
    CompoundFile.compress(str(src), str(dest))
    return dest


def test_truncated_file(tmp_path):
    dest = make_compound_file(tmp_path, {"big": os.urandom(100_000)})
    raw = dest.read_bytes()
    dest.write_bytes(raw[: len(raw) // 2])

    with pytest.raises(CompoundFileError):
        CompoundFile.decompress(str(dest), str(tmp_path / "out"))


def directory_offset(raw, name):
    """Return file offset of the directory entry named `name`."""

    pattern = struct.pack("<64sH", name.encode("utf-16-le"), len(name) * 2 + 2)
    assert raw.count(pattern) == 1
    return raw.index(pattern)


def test_directory_cycle(tmp_path):
    dest = make_compound_file(tmp_path, {"only": b"data"})
    raw = bytearray(dest.read_bytes())

    # Point the only child's left sibling back at itself
    root = directory_offset(raw, "Root Entry")
    (child,) = struct.unpack_from("<I", raw, root + 76)
    struct.pack_into("<I", raw, directory_offset(raw, "only") + 68, child)
    dest.write_bytes(bytes(raw))

    with pytest.raises(CompoundFileError, match="more than once"):
        CompoundFile.decompress(str(dest), str(tmp_path / "out"))


def test_invalid_start_sector(tmp_path):
    dest = make_compound_file(tmp_path, {"big": os.urandom(10_000)})
    raw = bytearray(dest.read_bytes())

    struct.pack_into("<I", raw, directory_offset(raw, "big") + 116, 0x00FFFFFF)
    dest.write_bytes(bytes(raw))

    with pytest.raises(CompoundFileError, match="sector number"):
        CompoundFile.decompress(str(dest), str(tmp_path / "out"))


def test_fat_chain_loop(tmp_path):
    # Directory chain is followed until ENDOFCHAIN, so a loop must be detected
    dest = make_compound_file(tmp_path, {"f": b"1"})
    raw = bytearray(dest.read_bytes())

    first_dir_sector, = struct.unpack_from("<I", raw, 48)
    first_fat_sector, = struct.unpack_from("<I", raw, 76)
    fat_offset = 512 + first_fat_sector * 512
    struct.pack_into("<I", raw, fat_offset + first_dir_sector * 4, first_dir_sector)
    dest.write_bytes(bytes(raw))

    with pytest.raises(CompoundFileError, match="loop"):
        CompoundFile.decompress(str(dest), str(tmp_path / "out"))
