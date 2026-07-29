import tempfile
import unittest
from pathlib import Path

from streamlit.testing.v1 import AppTest


ROOT = Path(__file__).resolve().parents[1]


class StreamlitAppTests(unittest.TestCase):
    def test_local_mode_processes_examples_without_moving_sources(self):
        before = len(list((ROOT / "exm").glob("*.pdf")))
        with tempfile.TemporaryDirectory() as temp_dir:
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=30)
            self.assertEqual(app.radio[0].value, "邮箱下载并处理")
            app.radio[0].set_value("本地文件夹处理").run(timeout=30)
            app.text_input[0].set_value(str(ROOT / "exm"))
            app.text_input[1].set_value(str(Path(temp_dir) / "output"))
            app.button[0].click().run(timeout=60)
            self.assertEqual(len(app.success), 1)
            self.assertTrue((Path(temp_dir) / "output" / "发票.xlsx").exists())
            self.assertTrue((Path(temp_dir) / "output" / "出租车报销单.xlsx").exists())
        self.assertEqual(len(list((ROOT / "exm").glob("*.pdf"))), before)


if __name__ == "__main__":
    unittest.main()
