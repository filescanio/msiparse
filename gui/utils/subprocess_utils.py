import subprocess
import sys
import shlex
from typing import List, Union, Optional, Any

def run_subprocess(
    args: Union[str, List[str]],
    check: bool = False,
    capture_output: bool = False,
    text: bool = False,
    input: Optional[Union[str, bytes]] = None,
    stdout: Optional[int] = None,
    stderr: Optional[int] = None,
    **kwargs: Any
) -> subprocess.CompletedProcess:
    """
    Runs a subprocess command with platform-specific handling.

    On Windows, it uses CREATE_NO_WINDOW to avoid opening a console window.
    Accepts arguments similar to subprocess.run.

    Args:
        args: The command arguments (list or string). If a string, it's split using shlex.
        check: If True, raises CalledProcessError on non-zero exit code.
        capture_output: If True, captures stdout and stderr.
        text: If True, decodes stdout/stderr as text.
        input: Data to send to the subprocess's stdin.
        stdout: Redirect stdout (e.g., subprocess.PIPE).
        stderr: Redirect stderr (e.g., subprocess.PIPE).
        **kwargs: Additional keyword arguments passed to subprocess.run.

    Returns:
        A subprocess.CompletedProcess object.
    """
    creationflags = 0
    if sys.platform == "win32":
        # Use getattr for safety in case CREATE_NO_WINDOW isn't defined (though highly unlikely)
        creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0) 

    # Ensure args is a list for subprocess.run
    if isinstance(args, str):
        args_list = shlex.split(args, posix=(sys.platform != "win32"))
    else:
        # Ensure it's a mutable list if passed as tuple, etc.
        args_list = list(args) 

    return subprocess.run(
        args_list,
        check=check,
        capture_output=capture_output,
        text=text,
        input=input,
        stdout=stdout,
        stderr=stderr,
        creationflags=creationflags,
        **kwargs
    ) 