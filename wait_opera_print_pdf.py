import os
import sys
import time
import glob
import shutil


def wait_for_opera_print_pdf(timeout: int, move_to_path: str = None) -> tuple:
    """
    Wait up to <timeout> seconds for an OperaPrint*.pdf to appear fully
    downloaded in the user's INetCache folder.

    Checks (recursively, case-insensitively):
        %LOCALAPPDATA%\\Microsoft\\Windows\\INetCache\\**\\OperaPrint*.pdf

    "Fully downloaded" means:
        1. File exists and size > 0
        2. Size unchanged between two stats 2 seconds apart
        3. File is not locked by another process (os.open O_RDWR succeeds)

    If move_to_path is provided (a full file path including filename), the
    verified PDF is moved there (overwriting if it exists) and the destination
    path is returned. On move failure, logs the error and returns (True, source_path).

    Returns:
        (True, path)  on success (destination path if moved, else source path)
        (False, None) on timeout
    Does not raise.
    """
    local_app_data = os.environ.get('LOCALAPPDATA') or os.path.join(
        os.environ.get('USERPROFILE', ''), 'AppData', 'Local')

    inet_cache = os.path.join(local_app_data, r'Microsoft\Windows\INetCache')

    if not os.path.isdir(inet_cache):
        print(f"INetCache folder not found: {inet_cache}")
        return False, None

    print(f"Waiting for OperaPrint*.pdf in INetCache: {inet_cache}")
    print(f"Timeout: {timeout} seconds")

    end_time = time.time() + timeout
    attempt = 1
    patterns = [
        os.path.join(inet_cache, "**", "OperaPrint*.pdf"),
        os.path.join(inet_cache, "**", "OPERAPRINT*.pdf"),
    ]

    # Track previously seen sizes so we only do the 2s stability check on
    # files whose size has already been observed once this run.
    seen_sizes = {}

    while time.time() < end_time:
        matches = []
        for pattern in patterns:
            matches.extend(glob.glob(pattern, recursive=True))

        if not matches:
            print(f"Attempt {attempt}: No OperaPrint*.pdf found yet.")
        else:
            print(f"Attempt {attempt}: Found {len(matches)} candidate(s).")

        for path in matches:
            try:
                size = os.path.getsize(path)
            except OSError as e:
                print(f"  Could not stat {path}: {e}")
                continue

            if size <= 0:
                print(f"  {os.path.basename(path)}: size 0, still downloading.")
                continue

            prev = seen_sizes.get(path)
            if prev is None:
                # First time seeing this file; record size and continue
                seen_sizes[path] = size
                print(f"  {os.path.basename(path)}: first sighting, size={size}. Will confirm stable next pass.")
                continue

            if size != prev:
                # Still growing; update and keep waiting
                seen_sizes[path] = size
                print(f"  {os.path.basename(path)}: size changed {prev} -> {size}, still writing.")
                continue

            # Size stable across two observations. Confirm not locked.
            print(f"  {os.path.basename(path)}: size stable at {size}. Checking write lock...")
            try:
                fd = os.open(path, os.O_RDWR)
                os.close(fd)
            except (PermissionError, OSError) as e:
                print(f"  {os.path.basename(path)}: locked by another process ({e}). Still downloading.")
                # Keep the seen size; will retry lock next pass.
                continue

            print(f"  {os.path.basename(path)}: fully downloaded at {path}")
            if move_to_path:
                try:
                    parent = os.path.dirname(move_to_path)
                    if parent:
                        os.makedirs(parent, exist_ok=True)
                    if os.path.exists(move_to_path):
                        try:
                            os.remove(move_to_path)
                        except OSError as e:
                            print(f"  Could not remove existing target {move_to_path}: {e}; shutil.move will try to replace.")
                    shutil.move(path, move_to_path)
                    print(f"  Moved to {move_to_path}")
                    return True, move_to_path
                except (OSError, shutil.Error) as e:
                    print(f"  Move failed: {e}. PDF remains at source: {path}")
                    return True, path
            return True, path

        attempt += 1
        remaining_time = end_time - time.time()
        if remaining_time > 0:
            print(f"Waiting 2 seconds before next attempt... ({remaining_time:.1f}s remaining)")
            time.sleep(2)

    print(f"Timeout reached after {timeout} seconds. No fully downloaded OperaPrint*.pdf found.")
    return False, None


if __name__ == "__main__":
    # CLI usage: python wait_opera_print_pdf.py <timeout> [move_to_path]
    if len(sys.argv) > 1:
        timeout_val = int(sys.argv[1])
    else:
        timeout_val = 60
    move_target = sys.argv[2] if len(sys.argv) > 2 else None
    success, file_path = wait_for_opera_print_pdf(timeout_val, move_target)
    if success:
        print(f"SUCCESS: {file_path}")
    else:
        print("FAILURE: no OperaPrint*.pdf downloaded within timeout.")