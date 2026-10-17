import os
import struct

import olefile
import pytest

from cfbpy import CompoundFile

import cfbcheck

# Names that cannot be created on a case-insensitive Windows filesystem
posix_only = pytest.mark.skipif(os.name == "nt", reason="not a valid Windows file name")


def make_tree(root, files, dirs=()):
    for path, data in files.items():
        full_path = os.path.join(root, *path.split("/"))
        os.makedirs(os.path.dirname(full_path), exist_ok=True)
        with open(full_path, "wb") as f:
            f.write(data)

    for path in dirs:
        os.makedirs(os.path.join(root, *path.split("/")), exist_ok=True)


def read_tree(root):
    """Return `{relative_path: bytes | None}` for a directory tree."""

    result = {}
    for dirpath, dirnames, filenames in os.walk(root):
        rel = os.path.relpath(dirpath, root)
        prefix = () if rel == "." else tuple(rel.split(os.sep))
        for name in dirnames:
            result["/".join(prefix + (name,))] = None
        for name in filenames:
            with open(os.path.join(dirpath, name), "rb") as f:
                result["/".join(prefix + (name,))] = f.read()
    return result


def expected_entries(files, dirs=()):
    """All paths, including implicit parent storages, in `read_tree` form."""

    result = {}
    for path in list(files) + list(dirs):
        parts = path.split("/")
        for depth in range(1, len(parts)):
            result["/".join(parts[:depth])] = None
    result.update({path: None for path in dirs})
    result.update(files)
    return result


def compress_and_check(tmp_path, files, dirs=()):
    """Compress a tree, then verify the output with two independent readers."""

    src = tmp_path / "src"
    src.mkdir()
    make_tree(src, files, dirs)
    dest = tmp_path / "out.cfb"

    CompoundFile.compress(str(src), str(dest))

    expected = expected_entries(files, dirs)

    # Strict structural validation
    parsed = cfbcheck.read(dest)
    assert {"/".join(path): data for path, data in parsed.items()} == expected

    # Cross-check with olefile, raising on anything it considers incorrect
    with olefile.OleFileIO(str(dest), raise_defects=olefile.DEFECT_INCORRECT) as ole:
        for path, data in files.items():
            assert ole.openstream(path).read() == data

    return dest


@pytest.mark.parametrize(
    "size",
    [1, 63, 64, 65, 511, 512, 513, 4095, 4096, 4097, 8192, 100_000],
)
def test_stream_sizes(tmp_path, size):
    compress_and_check(tmp_path, {"data": os.urandom(size)})


def test_empty_stream(tmp_path):
    compress_and_check(tmp_path, {"empty": b"", "other": b"x"})


def test_empty_source(tmp_path):
    compress_and_check(tmp_path, {})


def test_only_large_streams(tmp_path):
    # No mini stream at all
    compress_and_check(tmp_path, {"a": os.urandom(5000), "b": os.urandom(9000)})


def test_nested_storages(tmp_path):
    files = {
        "a.txt": b"hello",
        "sub/b.bin": b"x" * 5000,
        "sub/deep/c": b"y" * 100,
        "sub/deep/er/d": b"z" * 4096,
    }
    compress_and_check(tmp_path, files)


def test_deep_nesting(tmp_path):
    compress_and_check(tmp_path, {"/".join(["d"] * 30 + ["leaf"]): b"leaf"})


def test_empty_storage(tmp_path):
    compress_and_check(tmp_path, {"f": b"1"}, dirs=["empty", "nested/empty"])


def test_hidden_storage(tmp_path):
    compress_and_check(tmp_path, {".hidden/f": b"1", "..x": b"2"})


@pytest.mark.parametrize("count", [1, 2, 3, 4, 5, 7, 8, 15, 16, 17, 31, 100, 257])
def test_directory_tree(tmp_path, count):
    # cfbcheck verifies ordering and red-black invariants for each tree shape
    compress_and_check(tmp_path, {f"s/f{i}": bytes([i % 256]) for i in range(count)})


def test_name_ordering(tmp_path):
    # Shorter names sort first, then case-insensitive comparison
    names = ["b", "A", "ab", "AC", "z", "Zz", "가나다", "abc", "ABD", "_", "0"]
    compress_and_check(tmp_path, {name: name.encode() for name in names})


def test_unicode_names(tmp_path):
    compress_and_check(tmp_path, {"한글/스트림": b"data", "é": b"e"})


def test_many_streams(tmp_path):
    files = {f"d{i % 7}/f{i}": os.urandom(i % 300) for i in range(2000)}
    compress_and_check(tmp_path, files)


def test_large_stream_needs_difat(tmp_path):
    # More than 109 FAT sectors require DIFAT sectors
    compress_and_check(tmp_path, {"big": os.urandom(8 << 20), "small": b"s"})


def test_max_name_length(tmp_path):
    compress_and_check(tmp_path, {"n" * 31: b"ok"})


@pytest.mark.parametrize(
    "name",
    [
        "n" * 32,
        "a!b",
        pytest.param("a:b", marks=posix_only),
        pytest.param("a\\b", marks=posix_only),
    ],
)
def test_invalid_name(tmp_path, name):
    src = tmp_path / "src"
    make_tree(src, {name: b"x"})

    with pytest.raises(ValueError):
        CompoundFile.compress(str(src), str(tmp_path / "out.cfb"))


@posix_only
def test_case_insensitive_duplicate(tmp_path):
    src = tmp_path / "src"
    make_tree(src, {"Name": b"1", "nAME": b"2"})

    with pytest.raises(ValueError):
        CompoundFile.compress(str(src), str(tmp_path / "out.cfb"))


def test_default_destination(tmp_path):
    src = tmp_path / "src"
    make_tree(src, {"f": b"1"})

    CompoundFile.compress(str(src) + os.sep)

    assert cfbcheck.read(tmp_path / "src.cfb") == {("f",): b"1"}


def test_destination_inside_source(tmp_path):
    src = tmp_path / "src"
    make_tree(src, {"f": b"1"})

    CompoundFile.compress(str(src), str(src / "out.cfb"))

    assert cfbcheck.read(src / "out.cfb") == {("f",): b"1"}


def test_missing_source(tmp_path):
    with pytest.raises(NotADirectoryError):
        CompoundFile.compress(str(tmp_path / "missing"), str(tmp_path / "out.cfb"))

    assert not (tmp_path / "out.cfb").exists()


def test_failed_compress_removes_output(tmp_path):
    src = tmp_path / "src"
    make_tree(src, {"a": b"1", "n" * 32: b"2"})

    with pytest.raises(ValueError):
        CompoundFile.compress(str(src), str(tmp_path / "out.cfb"))

    assert not (tmp_path / "out.cfb").exists()


def test_deterministic_output(tmp_path):
    src = tmp_path / "src"
    make_tree(src, {f"d{i % 3}/f{i}": os.urandom(i * 97) for i in range(50)})

    CompoundFile.compress(str(src), str(tmp_path / "a.cfb"))
    CompoundFile.compress(str(src), str(tmp_path / "b.cfb"))

    assert (tmp_path / "a.cfb").read_bytes() == (tmp_path / "b.cfb").read_bytes()
