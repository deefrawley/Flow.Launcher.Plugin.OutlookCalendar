import sys
from pathlib import Path

root = Path(__file__).parent
lib_dir = root / "lib"
lib_dir.mkdir(parents=True, exist_ok=True)

sys.path.insert(0, root.as_posix())
sys.path.insert(0, lib_dir.as_posix())

from flogin.utils import setup_logging
from flogin import Pip

setup_logging()

with Pip(libs_dir=lib_dir) as pip:
    pip.ensure_installed("pywin32", module="pywintypes")

from OutlookAgendaPlugin.plugin import OutlookAgendaPlugin

if __name__ == "__main__":
    OutlookAgendaPlugin().run()
