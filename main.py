import os
import sys

import eel

from backend import app  # noqa: F401  # Ensures eel-exposed functions load.


def main() -> None:
    if getattr(sys, "frozen", False):
        # Running as a PyInstaller bundle — frontend is extracted to sys._MEIPASS
        frontend_path = os.path.join(sys._MEIPASS, "frontend")
    else:
        frontend_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "frontend")

    eel.init(frontend_path)
    eel.start(
        "index.html",
        size=(1920, 1080),
        position=(0, 0),
        block=True,
    )


if __name__ == "__main__":
    main()
