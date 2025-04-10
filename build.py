# build.py
import subprocess
import os
import shutil
import platform
import urllib.request
import tarfile
import sys

# --- Configuration ---
RUST_PROJECT_NAME = "msiparse"
PYTHON_GUI_DIR = "gui"
PYTHON_SPEC_FILE = "msiparse.spec"
PYTHON_DIST_NAME = "msiparse-gui"
ARTIFACT_DIR = "artifact"

WIN_7Z_URL = "https://www.7-zip.org/a/7zr.exe"
WIN_7Z_EXE = "7z.exe"
LINUX_7Z_URL = "https://www.7-zip.org/a/7z2409-linux-x64.tar.xz"
LINUX_7Z_ARCHIVE = "7z.tar.xz"
LINUX_7Z_EXE_IN_ARCHIVE = "7zzs" # Common name for static build in archive
LINUX_7Z_EXE = "7z"
# ---------------------

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
ARTIFACT_PATH = os.path.join(ROOT_DIR, ARTIFACT_DIR)

def run_command(cmd, cwd=None, check=True):
    """Runs a command in a subprocess."""
    print(f"Running: {' '.join(cmd)} {'in ' + cwd if cwd else ''}")
    process = subprocess.run(cmd, cwd=cwd, check=False, capture_output=True, text=True)
    if process.stdout:
        print("STDOUT:\n", process.stdout)
    if process.stderr:
        print("STDERR:\n", process.stderr, file=sys.stderr)
    if check and process.returncode != 0:
        print(f"Command failed with exit code {process.returncode}", file=sys.stderr)
        sys.exit(process.returncode)
    return process

def build_rust():
    """Builds the Rust project in release mode."""
    print("\n--- Building Rust Project ---")
    run_command(["cargo", "build", "--release"], cwd=ROOT_DIR)
    print("--- Rust Build Done ---")

def build_python_gui():
    """Installs dependencies and builds the Python GUI using PyInstaller."""
    print("\n--- Building Python GUI ---")
    gui_path = os.path.join(ROOT_DIR, PYTHON_GUI_DIR)
    # Install dependencies
    run_command([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"], cwd=gui_path)
    # Run PyInstaller
    run_command([sys.executable, "-m", "PyInstaller", PYTHON_SPEC_FILE, "--noconfirm", "--clean"], cwd=gui_path)
    print("--- Python GUI Build Done ---")

def prepare_artifacts():
    """Creates the artifact directory and moves build outputs."""
    print("\n--- Preparing Artifacts ---")
    if os.path.exists(ARTIFACT_PATH):
        print(f"Removing existing artifact directory: {ARTIFACT_PATH}")
        shutil.rmtree(ARTIFACT_PATH)
    os.makedirs(ARTIFACT_PATH)
    print(f"Created artifact directory: {ARTIFACT_PATH}")

    # Move Rust binary
    rust_exe_name = RUST_PROJECT_NAME
    if platform.system() == "Windows":
        rust_exe_name += ".exe"
    rust_src_path = os.path.join(ROOT_DIR, "target", "release", rust_exe_name)
    rust_dest_path = os.path.join(ARTIFACT_PATH, rust_exe_name)

    if os.path.exists(rust_src_path):
        print(f"Moving Rust binary: {rust_src_path} -> {rust_dest_path}")
        shutil.move(rust_src_path, rust_dest_path)
    else:
        print(f"Error: Rust binary not found at {rust_src_path}", file=sys.stderr)
        sys.exit(1)

    # Move Python GUI distribution
    # IMPORTANT: Adjust pyinstaller_dist_path if your spec file outputs differently
    #            (e.g., a single file instead of a directory, or a different name)
    pyinstaller_dist_path = os.path.join(ROOT_DIR, PYTHON_GUI_DIR, "dist", PYTHON_DIST_NAME)
    pyinstaller_dest_path = os.path.join(ARTIFACT_PATH, PYTHON_DIST_NAME)

    if os.path.isdir(pyinstaller_dist_path): # Assuming one-dir build
         print(f"Moving Python GUI dist: {pyinstaller_dist_path} -> {pyinstaller_dest_path}")
         shutil.move(pyinstaller_dist_path, pyinstaller_dest_path)
    elif os.path.isfile(pyinstaller_dist_path + ".exe"): # Check for one-file exe on Windows
         print(f"Moving Python GUI exe: {pyinstaller_dist_path}.exe -> {pyinstaller_dest_path}.exe")
         shutil.move(pyinstaller_dist_path + ".exe", pyinstaller_dest_path + ".exe")
    elif os.path.isfile(pyinstaller_dist_path): # Check for one-file exe on Linux/macOS
         print(f"Moving Python GUI executable: {pyinstaller_dist_path} -> {pyinstaller_dest_path}")
         shutil.move(pyinstaller_dist_path, pyinstaller_dest_path)
    else:
         print(f"Error: Python GUI distribution not found at {pyinstaller_dist_path} (checked dir, .exe, and file)", file=sys.stderr)
         sys.exit(1)

    print("--- Artifact Preparation Done ---")


def download_7z():
    """Downloads and prepares the 7-Zip executable based on the OS."""
    print("\n--- Downloading 7-Zip ---")
    system = platform.system()

    try:
        if system == "Windows":
            dest_path = os.path.join(ARTIFACT_PATH, WIN_7Z_EXE)
            print(f"Downloading {WIN_7Z_URL} to {dest_path}...")
            urllib.request.urlretrieve(WIN_7Z_URL, dest_path)
            print(f"{WIN_7Z_EXE} downloaded.")

        elif system == "Linux":
            archive_path = os.path.join(ARTIFACT_PATH, LINUX_7Z_ARCHIVE)
            dest_exe_path = os.path.join(ARTIFACT_PATH, LINUX_7Z_EXE)
            print(f"Downloading {LINUX_7Z_URL} to {archive_path}...")
            urllib.request.urlretrieve(LINUX_7Z_URL, archive_path)
            print(f"Extracting {LINUX_7Z_ARCHIVE}...")
            with tarfile.open(archive_path, "r:xz") as tar:
                # Extract only the target executable to avoid clutter
                extracted = False
                for member in tar.getmembers():
                    if os.path.basename(member.name) == LINUX_7Z_EXE_IN_ARCHIVE:
                         # Extract to artifact dir with the final desired name
                         member.name = LINUX_7Z_EXE
                         tar.extract(member, path=ARTIFACT_PATH)
                         print(f"Extracted {LINUX_7Z_EXE} to {ARTIFACT_PATH}")
                         extracted = True
                         break # Found what we needed
                if not extracted:
                     print(f"Error: Could not find {LINUX_7Z_EXE_IN_ARCHIVE} in {LINUX_7Z_ARCHIVE}", file=sys.stderr)
                     # Attempt generic extraction as fallback
                     print(f"Attempting generic extraction to {ARTIFACT_PATH}")
                     tar.extractall(path=ARTIFACT_PATH)


            print(f"Removing archive {archive_path}...")
            os.remove(archive_path)
            # Ensure the extracted binary is executable
            if os.path.exists(dest_exe_path):
                 os.chmod(dest_exe_path, 0o755)
                 print(f"Set {dest_exe_path} as executable.")
            else:
                 print(f"Error: Expected 7z executable not found at {dest_exe_path} after extraction.", file=sys.stderr)


        else:
            print(f"Warning: Unsupported OS for 7z download: {system}. Skipping.")

    except Exception as e:
        print(f"Error downloading or processing 7-Zip: {e}", file=sys.stderr)
        # Decide if this should be fatal? For now, just warn.
        # sys.exit(1)

    print("--- 7-Zip Download Done ---")


if __name__ == "__main__":
    build_rust()
    build_python_gui()
    prepare_artifacts()
    download_7z()
    print("\nBuild process completed.")
    print(f"Artifacts are available in: {ARTIFACT_PATH}") 