# cfbpy

`cfbpy` is a Python package that allows you to compress and decompress compound file binary formats,
similar to how zip/unzip works for regular zip files.
It provides an easy-to-use interface for handling compound files through the CompoundFile class.

## Features

- Compress a directory into a compound file format.
- Decompress a compound file back into a directory.

## Installation

You can install cfbpy via pip:

```bash
pip install cfbpy
```

## Usage

Here’s how you can use `cfbpy` to compress and decompress files:

### Compress a Directory

```python
from cfbpy import compress

compress('path/to/directory', 'path/to/output.cfb')
```

### Decompress a Compound File

```python
from cfbpy import decompress

decompress('path/to/input.cfb', 'path/to/extracted_directory')
```

The same functions are also available as `CompoundFile.compress` and `CompoundFile.decompress`.

## Methods

- `compress(src, dest=None)`: Compresses the contents of the `src` directory into a compound file at `dest` (default: `<src>.cfb`).
- `decompress(src, dest=None)`: Decompresses the compound file from `src` to the directory `dest` (default: `src` without its extension).

Compound file names are limited by the format, so `compress` raises `ValueError` when a file or directory name

- is longer than 31 UTF-16 code units,
- contains any of `/ \ : !`, or
- differs from a sibling only in letter case.

`decompress` raises `ValueError` for entry names that would escape the destination directory, such as `..`.

Files are written as a major version 3 compound file (512-byte sectors).

## Development

```bash
pip install -e ".[test]"
pytest                       # run tests
python benchmarks/bench.py   # measure compress/decompress time and memory
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change.
