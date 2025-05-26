import pytest
import os
import subprocess
import sys

# --------- Add gui directory to Python path for test discovery/execution ---------
_WORKSPACE_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
_GUI_DIR = os.path.join(_WORKSPACE_ROOT, 'gui')
if _GUI_DIR not in sys.path:
    print(f"\nINFO: Inserting GUI directory into sys.path: {_GUI_DIR}")
    sys.path.insert(0, _GUI_DIR)
# ---------------------------------------------------------------------------------

@pytest.fixture(scope='session', autouse=True)
def setup_test_environment(request):
    """
    Fixture to ensure the msiparse executable is built before tests run.
    Runs once per test session automatically.
    """
    print("\nSetting up test environment (running cargo build)...")
    workspace_root = _WORKSPACE_ROOT

    try:
        cargo_command = ["cargo", "build", "--release"]
        result = subprocess.run(
            cargo_command,
            cwd=workspace_root,
            capture_output=True,
            text=True,
            check=False
        )

        if result.returncode != 0:
            pytest.fail("Cargo build failed, cannot proceed with tests.")
    except Exception as e:
        pytest.fail(f"Error during test environment setup: {e}")

    print("Test environment setup finished successfully.")
    yield
