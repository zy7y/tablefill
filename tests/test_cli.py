import os
from pathlib import Path
from unittest import TestCase


class TestCli(TestCase):
    def test_cli_generate(self):
        os.system(
            r'Python ../tablefill/cli.py generate "E:\coding\tablefill\examples\demo.json" "E:\coding\tablefill\examples\demo.xlsx" demo.xls '
        )
        p = Path.cwd().joinpath("demo.xls")
        assert p.exists()
