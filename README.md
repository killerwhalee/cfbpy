# cfbpy

`cfbpy` is a Python package that allows you to compress and decompress compound file binary formats,
similar to how zip/unzip works for regular zip files.
It provides an easy-to-use interface for handling compound files through the CompoundFile class.

## Features

Compress a directory into a compound file format.
Decompress a compound file back into a directory.
Installation
You can install cfbpy via pip:

```bash
pip install cfbpy
```

## Usage

Here’s how you can use `cfbpy` to compress and decompress files:

### Compress a Directory

```python
from cfbpy import CompoundFile

cf = CompoundFile()
cf.compress('path/to/directory', 'path/to/output.cfb')
```

### Decompress a Compound File

```python
from cfbpy import CompoundFile

cf = CompoundFile()
cf.decompress('path/to/input.cfb', 'path/to/extracted_directory')
```

## Methods

- compress(src, dest): Compresses the contents of the `src` directory into a compound file at `dest`.
- decompress(src, dest): Decompresses the compound file from `src` to the directory `dest`.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change.
