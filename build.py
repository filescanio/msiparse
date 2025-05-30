# build.py
import subprocess
import os
import shutil
import platform
import urllib.request
import tarfile
import sys
import argparse

# --- Configuration ---
RUST_PROJECT_NAME = "msiparse"
PYTHON_GUI_DIR = "gui"
PYTHON_SPEC_FILE = "msiparse.spec"
PYTHON_DIST_NAME = "msiparse-gui"
ARTIFACT_DIR = "artifact"

WIN_7Z_URL = "https://www.7-zip.org/a/7z2409-extra.7z"
WIN_7Z_ARCHIVE_NAME = "7z-win-extra.7z"
WIN_7Z_EXTRACT_EXE = "7za.exe"
WIN_7ZR_DOWNLOAD_URL = "https://www.7-zip.org/a/7zr.exe"
WIN_7ZR_TEMP_NAME = "7zr_temp.exe"
LINUX_7Z_URL = "https://www.7-zip.org/a/7z2409-linux-x64.tar.xz"
LINUX_7Z_ARCHIVE = "7z-linux.tar.xz"
LINUX_7Z_EXE_IN_ARCHIVE = "7zzs"
LINUX_7Z_EXE = "7z"
MACOS_7Z_URL = "https://www.7-zip.org/a/7z2409-mac.tar.xz"
MACOS_7Z_ARCHIVE = "7z-mac.tar.xz"
MACOS_7Z_EXE_IN_ARCHIVE = "7zz"
MACOS_7Z_EXE = "7z"
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
    run_command(["cargo", "clean"], cwd=ROOT_DIR)
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
    pyinstaller_dist_dir = os.path.join(ROOT_DIR, PYTHON_GUI_DIR, "dist")
    pyinstaller_dist_name = None

    # Find the actual output name from PyInstaller (could be dir or file)
    if os.path.isdir(pyinstaller_dist_dir):
        items = os.listdir(pyinstaller_dist_dir)
        if items: # Check if dist directory is not empty
             # Prefer directory if it exists and matches expected name, otherwise take first item
             if PYTHON_DIST_NAME in items and os.path.isdir(os.path.join(pyinstaller_dist_dir, PYTHON_DIST_NAME)):
                  pyinstaller_dist_name = PYTHON_DIST_NAME
             else:
                  # Might be a single file executable or differently named directory
                  pyinstaller_dist_name = items[0]

    if pyinstaller_dist_name:
        pyinstaller_src_path = os.path.join(pyinstaller_dist_dir, pyinstaller_dist_name)
        pyinstaller_dest_path = os.path.join(ARTIFACT_PATH, pyinstaller_dist_name)
        print(f"Moving Python GUI distribution: {pyinstaller_src_path} -> {pyinstaller_dest_path}")
        shutil.move(pyinstaller_src_path, pyinstaller_dest_path)
    else:
        print(f"Error: Python GUI distribution not found in {pyinstaller_dist_dir}", file=sys.stderr)
        sys.exit(1)

    print("--- Artifact Preparation Done ---")


def download_7z():
    """Downloads and prepares the 7-Zip executable based on the OS."""
    print("\n--- Downloading 7-Zip ---")
    system = platform.system()
    success = False # Track overall success
    final_exe_path = None # Store the final path of the downloaded executable

    try:
        if system == "Windows":
            # Need 7zr.exe to extract the .7z archive containing 7za.exe
            temp_7zr_path = os.path.join(ARTIFACT_PATH, WIN_7ZR_TEMP_NAME)
            archive_path = os.path.join(ARTIFACT_PATH, WIN_7Z_ARCHIVE_NAME)
            final_exe_path = os.path.join(ARTIFACT_PATH, WIN_7Z_EXTRACT_EXE)

            print(f"Downloading temporary extractor {WIN_7ZR_TEMP_NAME}...")
            urllib.request.urlretrieve(WIN_7ZR_DOWNLOAD_URL, temp_7zr_path)

            print(f"Downloading main archive {WIN_7Z_ARCHIVE_NAME}...")
            urllib.request.urlretrieve(WIN_7Z_URL, archive_path)

            print(f"Extracting {WIN_7Z_EXTRACT_EXE} using {WIN_7ZR_TEMP_NAME}...")
            # Command: 7zr_temp.exe e archive.7z -o<output_dir> <file_to_extract> -y
            extract_cmd = [
                temp_7zr_path,
                "e",
                archive_path,
                f"-o{ARTIFACT_PATH}",
                WIN_7Z_EXTRACT_EXE,
                "-y"
            ]
            run_command(extract_cmd, check=True) # Check=True will raise on failure

            if os.path.exists(final_exe_path):
                success = True
            else:
                print(f"Error: Failed to extract {WIN_7Z_EXTRACT_EXE} from {WIN_7Z_ARCHIVE_NAME}", file=sys.stderr)

            # Cleanup temporary files
            print("Cleaning up temporary download files...")
            if os.path.exists(temp_7zr_path):
                try:
                    os.remove(temp_7zr_path)
                except OSError as e:
                    print(f"Warning: Could not remove temporary extractor {temp_7zr_path}: {e}", file=sys.stderr)
            if os.path.exists(archive_path):
                try:
                    os.remove(archive_path)
                except OSError as e:
                    print(f"Warning: Could not remove temporary archive {archive_path}: {e}", file=sys.stderr)


        elif system in ("Linux", "Darwin"): # Handle Linux and macOS together
            if system == "Linux":
                archive_url = LINUX_7Z_URL
                archive_filename = LINUX_7Z_ARCHIVE
                exe_in_archive = LINUX_7Z_EXE_IN_ARCHIVE
                final_exe_name = LINUX_7Z_EXE
            else: # system == "Darwin"
                archive_url = MACOS_7Z_URL
                archive_filename = MACOS_7Z_ARCHIVE
                exe_in_archive = MACOS_7Z_EXE_IN_ARCHIVE
                final_exe_name = MACOS_7Z_EXE

            archive_path = os.path.join(ARTIFACT_PATH, archive_filename)
            final_exe_path = os.path.join(ARTIFACT_PATH, final_exe_name) # Assign final_exe_path

            print(f"Downloading archive {archive_filename}...") # Added print
            urllib.request.urlretrieve(archive_url, archive_path)

            print(f"Extracting {final_exe_name} from {archive_filename}...") # Added print
            with tarfile.open(archive_path, "r:xz") as tar:
                extracted = False
                # Attempt direct extraction of the target executable
                for member in tar.getmembers():
                    if os.path.basename(member.name) == exe_in_archive:
                        member.name = final_exe_name # Rename on extraction
                        tar.extract(member, path=ARTIFACT_PATH)
                        extracted = True
                        print(f"Extracted {final_exe_name} directly.") # Added print
                        break

                # Fallback: Generic extraction if direct failed
                if not extracted:
                    print(f"Could not find {exe_in_archive} directly in archive, attempting generic extraction...")
                    tar.extractall(path=ARTIFACT_PATH)
                    found_generic = False
                    # Search for the executable in extracted files/dirs
                    for item in os.listdir(ARTIFACT_PATH):
                        item_path = os.path.join(ARTIFACT_PATH, item)
                        potential_exe_path = None

                        if os.path.isfile(item_path) and item == exe_in_archive:
                            potential_exe_path = item_path
                        elif os.path.isdir(item_path):
                            # Check inside a potential subdirectory
                            path_in_subdir = os.path.join(item_path, exe_in_archive)
                            if os.path.isfile(path_in_subdir):
                                potential_exe_path = path_in_subdir

                        if potential_exe_path:
                            print(f"Found {exe_in_archive} at {potential_exe_path} after generic extraction.") # Added print
                            # Move/rename to the final destination path
                            if potential_exe_path != final_exe_path:
                                print(f"Moving {potential_exe_path} to {final_exe_path}") # Added print
                                shutil.move(potential_exe_path, final_exe_path)
                            else:
                                # Already in the right place, possibly renamed if item==exe_in_archive
                                pass
                            found_generic = True

                            # Attempt cleanup of extracted dir if applicable
                            if os.path.isdir(item_path) and item_path != ARTIFACT_PATH: # Avoid deleting artifact dir itself
                                try:
                                    print(f"Cleaning up temporary directory {item_path}") # Added print
                                    shutil.rmtree(item_path)
                                except OSError as e:
                                    print(f"Warning: Could not remove temporary extraction dir {item_path}: {e}", file=sys.stderr)
                            break # Found it, stop searching

                    if not found_generic:
                        print(f"Warning: Could not find {exe_in_archive} even after generic extraction.", file=sys.stderr)
                    else:
                        extracted = True # Mark success if found via fallback

            # Cleanup and set permissions
            print(f"Cleaning up archive {archive_path}") # Added print
            os.remove(archive_path)
            if extracted and os.path.exists(final_exe_path):
                print(f"Setting execute permission on {final_exe_path}") # Added print
                os.chmod(final_exe_path, 0o755)
                success = True
            elif extracted:
                # This case means extraction happened but the final file isn't there
                print(f"Error: Expected 7z executable not found at {final_exe_path} after extraction steps.", file=sys.stderr)

        else:
            print(f"Warning: Unsupported OS for 7z download: {system}. Skipping.")

    except Exception as e:
        print(f"Error downloading or processing 7-Zip: {e}", file=sys.stderr)
        # Optional: Decide if lack of 7z should halt the build
        # sys.exit(1)

    # Print a summary status message
    if success and final_exe_path:
        print(f"7-Zip prepared successfully: {final_exe_path}")
    elif success:
        # Should not happen if logic is correct, but handle gracefully
        print(f"7-Zip prepared successfully for {system} (path variable not set).")
    else:
        print(f"7-Zip preparation failed or was skipped for {system}.")

    print("--- 7-Zip Download Done ---")


if __name__ == "__main__":
    # --- Argument Parsing ---
    parser = argparse.ArgumentParser(description="Build script for msiparse CLI and GUI.")
    parser.add_argument(
        "--cli-only",
        action="store_true",
        help="Build only the Rust CLI application."
    )
    args = parser.parse_args()
    # ------------------------

    build_rust()

    if not args.cli_only:
        # These steps are only needed for the full bundle
        build_python_gui()
        prepare_artifacts()
        download_7z()
        print("Full build process completed.")
        print(f"Artifacts are available in: {ARTIFACT_PATH}")
    else:
        # Only the rust build was performed
        rust_exe_name = RUST_PROJECT_NAME
        if platform.system() == "Windows":
            rust_exe_name += ".exe"
        rust_bin_path = os.path.join(ROOT_DIR, "target", "release", rust_exe_name)
        print("CLI-only build process completed.")
        print(f"Rust binary available at: {rust_bin_path}") 
