from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from src.reporting import write_json


class SerializationTests(unittest.TestCase):
    def test_unsupported_value_raises_native_type_error(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "payload.json"

            with self.assertRaises(TypeError):
                write_json(path, {"value": object()})

            self.assertFalse(path.exists())


if __name__ == "__main__":
    unittest.main()
