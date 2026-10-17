"""
Benchmark compress/decompress on generated directory trees.

Usage: python benchmarks/bench.py [--sizes 1 4 16 64] [--files 20000] [--no-memory]

Time and peak memory are measured in separate runs, as tracing allocations slows down the code under test.

"""

import argparse
import os
import shutil
import sys
import tempfile
import time
import tracemalloc

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))

from cfbpy import CompoundFile  # noqa: E402


def make_large(root, size_mb):
    os.makedirs(root)
    with open(os.path.join(root, "big"), "wb") as f:
        f.write(os.urandom(size_mb << 20))


def make_many(root, count, dirs=20):
    for i in range(count):
        storage = os.path.join(root, f"d{i % dirs}")
        os.makedirs(storage, exist_ok=True)
        with open(os.path.join(storage, f"file_{i}"), "wb") as f:
            f.write(b"q" * 200)


def measure(func, args, cleanup, memory):
    start = time.perf_counter()
    try:
        func(*args)
    except Exception as e:
        return None, None, f"{type(e).__name__}: {e}"
    elapsed = time.perf_counter() - start

    if not memory:
        return elapsed, None, None

    cleanup()
    tracemalloc.start()
    try:
        func(*args)
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()

    return elapsed, peak, None


def run(label, workdir, build, memory):
    src = os.path.join(workdir, label)
    dest = src + ".cfb"
    out = src + ".out"
    build(src)

    rows = []
    for name, func, args, cleanup in (
        ("compress", CompoundFile.compress, (src, dest), lambda: os.remove(dest)),
        ("decompress", CompoundFile.decompress, (dest, out), lambda: shutil.rmtree(out)),
    ):
        elapsed, peak, error = measure(func, args, cleanup, memory)
        rows.append((label, name, elapsed, peak, error))
        if error:
            break

    for path in (src, out):
        shutil.rmtree(path, ignore_errors=True)
    if os.path.exists(dest):
        os.remove(dest)

    return rows


def main():
    parser = argparse.ArgumentParser(description=__doc__.strip().splitlines()[0])
    parser.add_argument("--sizes", type=int, nargs="*", default=[1, 4, 16, 64], help="single-stream sizes in MB")
    parser.add_argument("--files", type=int, default=20000, help="number of small streams")
    parser.add_argument("--no-memory", dest="memory", action="store_false", help="skip peak memory measurement")
    args = parser.parse_args()

    cases = [(f"large_{mb}MB", lambda path, mb=mb: make_large(path, mb)) for mb in args.sizes]
    if args.files:
        cases.append((f"small_x{args.files}", lambda path: make_many(path, args.files)))

    print(f"{'case':<16} {'op':<11} {'time (s)':>9} {'peak mem (MB)':>14}")
    with tempfile.TemporaryDirectory() as workdir:
        for label, build in cases:
            for case, op, elapsed, peak, error in run(label, workdir, build, args.memory):
                if error:
                    print(f"{case:<16} {op:<11} FAILED: {error}", flush=True)
                    continue

                memory = f"{peak / (1 << 20):>14.1f}" if peak is not None else f"{'-':>14}"
                print(f"{case:<16} {op:<11} {elapsed:>9.3f} {memory}", flush=True)


if __name__ == "__main__":
    main()
