#!/usr/bin/env python3
"""Edge Cleaner Diagnostic - read-only.

Run this on each machine and compare the outputs side by side.
It does NOT modify anything (registry, files, or processes). It only
reads state and probes access permissions to explain why the Edge
cache/history cleaner behaves differently across machines.
"""

import json
import os
import winreg
import psutil


def sep(title):
    print("\n" + "=" * 70)
    print(title)
    print("=" * 70)


def yn(b):
    return "YES" if b else "NO"


def safe_json_read(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f), None
    except FileNotFoundError:
        return None, "file not found"
    except PermissionError as e:
        return None, f"permission denied: {e}"
    except json.JSONDecodeError as e:
        return None, f"JSON parse error: {e}"
    except Exception as e:
        return None, f"{type(e).__name__}: {e}"


def file_writable(path):
    if not os.path.isfile(path):
        return "n/a (file missing)"
    try:
        fd = os.open(path, os.O_WRONLY | os.O_APPEND)
        os.close(fd)
        return "YES"
    except PermissionError:
        return "NO (locked / denied)"
    except Exception as e:
        return f"NO ({type(e).__name__}: {e})"


# --- Sections ---------------------------------------------------------------

def section_environment():
    sep("[1] Environment")
    print(f"USERPROFILE    = {os.environ.get('USERPROFILE', '<unset>')}")
    print(f"LOCALAPPDATA   = {os.environ.get('LOCALAPPDATA', '<unset>')}")
    print(f"USERNAME       = {os.environ.get('USERNAME', '<unset>')}")
    print(f"COMPUTERNAME   = {os.environ.get('COMPUTERNAME', '<unset>')}")
    try:
        import ctypes
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        is_admin = "?"
    print(f"Running as Admin = {is_admin}")
    print(f"Python arch     = {('64-bit' if winreg.HKEY_CURRENT_USER else '?')}")


def section_processes():
    sep("[2] Edge Processes (is anything running?)")
    edge_names = {'msedge.exe', 'microsoftedge.exe', 'microsoftedgecp.exe', 'microsoftedgesh.exe'}
    found = []
    for proc in psutil.process_iter(['name', 'pid', 'create_time']):
        try:
            n = (proc.info['name'] or '').lower()
            if n in edge_names:
                found.append((n, proc.info['pid']))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    print(f"Edge processes running: {len(found)}")
    for n, pid in found:
        print(f"  - {n} (PID {pid})")
    if found:
        print("NOTE: Background msedge.exe = Startup Boost is likely ON.")
        print("      Startup Boost can respawn Edge mid-cleaning and lock Preferences.")


def section_data_dirs():
    sep("[3] Edge Data Folder Detection (the #1 cross-machine difference)")
    local_app_data = os.environ.get('LOCALAPPDATA') or os.path.join(
        os.environ.get('USERPROFILE', ''), 'AppData', 'Local')

    print("UserDataDir policy values (custom data folder overrides):")
    policy_keys = [
        ("HKLM",            winreg.HKEY_LOCAL_MACHINE, r'Software\Policies\Microsoft\Edge'),
        ("HKLM WOW6432Node", winreg.HKEY_LOCAL_MACHINE, r'Software\WOW6432Node\Policies\Microsoft\Edge'),
        ("HKCU",            winreg.HKEY_CURRENT_USER,  r'Software\Policies\Microsoft\Edge'),
    ]
    policy_dirs = []
    for label, root, path in policy_keys:
        try:
            with winreg.OpenKey(root, path) as key:
                val, _ = winreg.QueryValueEx(key, 'UserDataDir')
                print(f"  {label:18} = {val!r}")
                if val:
                    policy_dirs.append(os.path.expandvars(val))
        except (FileNotFoundError, OSError):
            print(f"  {label:18} = <not set / key missing>")

    print("\nDefault per-channel locations:")
    channel_paths = [
        ("Stable",  os.path.join(local_app_data, r'Microsoft\Edge\User Data')),
        ("Beta",    os.path.join(local_app_data, r'Microsoft\Edge Beta\User Data')),
        ("Dev",     os.path.join(local_app_data, r'Microsoft\Edge Dev\User Data')),
        ("Canary",  os.path.join(local_app_data, r'Microsoft\Edge SxS\User Data')),
    ]
    found_dirs = []
    for label, p in channel_paths:
        exists = os.path.isdir(p)
        print(f"  {label:7} {p}  [{'EXISTS' if exists else 'missing'}]")
        if exists:
            found_dirs.append(p)
    for d in policy_dirs:
        if d not in found_dirs and os.path.isdir(d):
            print(f"  Policy  {d}  [EXISTS]")
            found_dirs.append(d)

    print(f"\n=> Detected data dirs: {len(found_dirs)}")
    for d in found_dirs:
        print(f"   - {d}")
    return found_dirs


def section_profiles(found_dirs):
    sep("[4] Profiles & Preferences per Data Dir (where restore dialog comes from)")
    if not found_dirs:
        print("No data dirs detected - nothing to inspect.")
        return []

    all_profile_reports = []
    for data_dir in found_dirs:
        channel = os.path.basename(os.path.dirname(data_dir))
        print(f"\n--- Channel: {channel} ---")
        print(f"Path: {data_dir}")

        # Local State (has profile.info_cache with exit_type per profile)
        local_state_path = os.path.join(data_dir, 'Local State')
        local_state, err = safe_json_read(local_state_path)
        if err:
            print(f"  Local State: {err}")
        else:
            info_cache = local_state.get('profile', {}).get('info_cache', {}) if local_state else {}
            if info_cache:
                print("  Local State profile exit_types:")
                for prof_name, info in info_cache.items():
                    print(f"    {prof_name}: exit_type={info.get('exit_type', '<missing>')}, "
                          f"exited_cleanly={info.get('exited_cleanly', '<missing>')}")

        # Per-profile directories
        profile_dirs = []
        try:
            for name in os.listdir(data_dir):
                full = os.path.join(data_dir, name)
                if os.path.isdir(full) and (name == 'Default' or name.startswith('Profile ')):
                    profile_dirs.append((name, full))
        except Exception as e:
            print(f"  Could not list {data_dir}: {e}")
            continue

        print(f"  Profile dirs found: {len(profile_dirs)}")
        for name, pdir in profile_dirs:
            prefs_path = os.path.join(pdir, 'Preferences')
            last_tabs = os.path.join(pdir, 'Last Tabs')
            last_session = os.path.join(pdir, 'Last Session')
            print(f"\n  [{name}]")
            print(f"    Preferences exists : {os.path.isfile(prefs_path)}")
            print(f"    Preferences writable: {file_writable(prefs_path)}")
            print(f"    Last Tabs exists  : {os.path.exists(last_tabs)}")
            print(f"    Last Session exist : {os.path.exists(last_session)}")
            prefs, perr = safe_json_read(prefs_path)
            if prefs is not None:
                prof = prefs.get('profile', {})
                print(f"    profile.exit_type        = {prof.get('exit_type', '<missing>')!r}")
                print(f"    profile.exited_cleanly   = {prof.get('exited_cleanly', '<missing>')!r}")
            else:
                print(f"    Preferences read error: {perr}")
            all_profile_reports.append((channel, name, prefs_path, perr))
    return all_profile_reports


def section_registry_access():
    sep("[5] Registry Policy Access (HideRestoreDialogEnabled)")
    value_name = 'HideRestoreDialogEnabled'
    tests = [
        ("HKCU", winreg.HKEY_CURRENT_USER,  r'Software\Policies\Microsoft\Edge'),
        ("HKLM", winreg.HKEY_LOCAL_MACHINE, r'Software\Policies\Microsoft\Edge'),
        ("HKLM WOW64", winreg.HKEY_LOCAL_MACHINE, r'Software\WOW6432Node\Policies\Microsoft\Edge'),
    ]
    for label, root, path in tests:
        print(f"\n  {label}\\{path}")
        # Read current value if present
        try:
            with winreg.OpenKey(root, path) as key:
                try:
                    val, _ = winreg.QueryValueEx(key, value_name)
                    print(f"    Current value: {val}")
                except FileNotFoundError:
                    print(f"    Current value: <not set>")
                # Probe KEY_WRITE on existing key
                try:
                    with winreg.OpenKey(root, path, 0, winreg.KEY_WRITE) as k:
                        print(f"    Open with KEY_WRITE: OK (can modify)")
                except PermissionError:
                    print(f"    Open with KEY_WRITE: DENIED (would fail to write)")
                except FileNotFoundError:
                    print(f"    Open with KEY_WRITE: key disappeared")
        except FileNotFoundError:
            print(f"    Key does not exist")
            # Try to test if we can CREATE the leaf by safe create+delete of a temp child
            # We must walk down; if any level missing, we report inability to probe creation safely.
            print(f"    Create ability: not probed (would require creating parent keys)")


def section_startup_boost(found_dirs):
    sep("[6] Startup Boost Setting")
    for data_dir in found_dirs:
        channel = os.path.basename(os.path.dirname(data_dir))
        local_state_path = os.path.join(data_dir, 'Local State')
        ls, err = safe_json_read(local_state_path)
        if err:
            print(f"  {channel}: Local State {err}")
            continue
        boost_enabled = ls.get('browser', {}).get('enabled_labs', []) if ls else []
        # Startup Boost lives under browser.startup_boost_enabled in Local State (Chromium-ish)
        startup_boost = ls.get('browser', {}).get('startup_boost_enabled', '<missing>')
        print(f"  {channel}: startup_boost_enabled = {startup_boost!r}")
        # Also check the user-level setting under root
        root_boost = ls.get('startup_boost_enabled', '<missing>')
        if root_boost != '<missing>':
            print(f"  {channel}: root startup_boost_enabled = {root_boost!r}")


def section_summary(found_dirs, profile_reports):
    sep("[7] Summary - Likely Causes If Cleaner Fails Here")
    if not found_dirs:
        print("  - No Edge data folder detected. UserDataDir policy may point elsewhere,")
        print("    or Edge has never been run on this machine.")
    if not profile_reports:
        print("  - No readable Preferences found. Restore suppression cannot work.")
    admin = False
    try:
        import ctypes
        admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        pass
    if not admin:
        print("  - Not running as admin. HKLM registry writes will fail.")
    # Check if any Preferences is unwritable
    unwritable = [p for (_, _, p, _) in profile_reports if file_writable(p).startswith("NO")]
    if unwritable:
        print(f"  - {len(unwritable)} Preferences file(s) are NOT writable:")
        for p in unwritable:
            print(f"      {p}")
        print("    Cause: Edge process still running (Startup Boost) or AV holding the file.")


def main():
    print("Edge Cleaner Diagnostic (READ-ONLY - changes nothing)")
    print(f"Run on both machines and compare outputs.")
    try:
        section_environment()
        section_processes()
        found_dirs = section_data_dirs()
        profile_reports = section_profiles(found_dirs)
        section_registry_access()
        section_startup_boost(found_dirs)
        section_summary(found_dirs, profile_reports)
        print("\n" + "=" * 70)
        print("Diagnostic complete. Compare this output between the two machines.")
        print("=" * 70)
    except Exception as e:
        import traceback
        print(f"\nDiagnostic failed: {e}")
        print(traceback.format_exc())


if __name__ == "__main__":
    main()