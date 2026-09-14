import os
import sys
import tempfile

import officecli


if not sys.platform.startswith("win"):
    previous = os.environ.get("TMPDIR")
    try:
        os.environ["TMPDIR"] = "/" + ("nested/" * 20) + "tmp"
        main, ping = officecli.pipe_paths("/tmp/officecli-pipe-test.xlsx")
        expected_dir = f"/tmp/officecli-{os.getuid()}"
        assert os.path.dirname(main) == expected_dir
        assert ping == main + "-ping"
    finally:
        if previous is None:
            os.environ.pop("TMPDIR", None)
        else:
            os.environ["TMPDIR"] = previous

if sys.platform == "darwin":
    with tempfile.TemporaryDirectory(prefix="officecli-pipe-real-") as real_dir:
        link_dir = real_dir + "-link"
        real_file = os.path.join(real_dir, "book.xlsx")
        open(real_file, "w", encoding="utf-8").close()
        try:
            os.symlink(real_dir, link_dir)
            assert officecli.pipe_paths(os.path.join(link_dir, "book.xlsx")) == officecli.pipe_paths(real_file)
        finally:
            os.unlink(link_dir)

print("python pipe-path test PASS")
