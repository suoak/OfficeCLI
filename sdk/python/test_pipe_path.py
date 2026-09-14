import os
import sys

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

print("python pipe-path test PASS")
