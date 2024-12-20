import PyInstaller.__main__
import os

# Get the absolute path of the current directory
current_dir = os.path.abspath(os.path.dirname(__file__))

PyInstaller.__main__.run(
    [
        "main.py",
        "--name=Splitter",
        "--onedir",
        "--windowed",
        "--icon=logo.ico",
        "--add-data=logo.ico;.",
        "--clean",
        "--noconfirm",
    ]
)
