#!/usr/bin/env python3
"""Find where Edge policies are actually saved in the registry.

Edge reads policies from FOUR registry trees (mandatory + recommended,
each in HKLM and HKCU), plus their WOW6432Node mirrors. This script
scans all of them and prints every value found, with the full path,
so you can see exactly where edge://policy is pulling from.

It also scans UiPath PIP Browser Profiles and DualEngine custom Edge
data directories to confirm whether those alternate Edge profiles
exist and what restore-related state they hold.

READ-ONLY - changes nothing.
"""

import json
import os
import winreg


# All locations Edge reads policies from, in priority order.
# "Mandatory" = under ...\Policies\Microsoft\Edge (user cannot override)
# "Recommended" = under ...\Microsoft\Edge (user can override in settings)
EDGE_POLICY_LOCATIONS = [
    # (label, root, subpath, type)
    ("HKLM Mandatory",      winreg.HKEY_LOCAL_MACHINE, r'Software\Policies\Microsoft\Edge',                'mandatory'),
    ("HKLM WOW64 Mandatory", winreg.HKEY_LOCAL_MACHINE, r'Software\WOW6432Node\Policies\Microsoft\Edge',   'mandatory'),
    ("HKCU Mandatory",      winreg.HKEY_CURRENT_USER,  r'Software\Policies\Microsoft\Edge',                'mandatory'),
    ("HKLM Recommended",    winreg.HKEY_LOCAL_MACHINE, r'Software\Microsoft\Edge',                          'recommended'),
    ("HKLM WOW64 Recommended", winreg.HKEY_LOCAL_MACHINE, r'Software\WOW6432Node\Microsoft\Edge',          'recommended'),
    ("HKCU Recommended",    winreg.HKEY_CURRENT_USER,  r'Software\Microsoft\Edge',                          'recommended'),
]


def enum_key_values(root, path):
    """Return list of (value_name, value_data, value_type_name) for a key, or None if key missing."""
    try:
        key = winreg.OpenKey(root, path)
    except FileNotFoundError:
        return None
    except PermissionError:
        return "ACCESS_DENIED"
    except OSError as e:
        return f"ERROR: {e}"

    results = []
    try:
        i = 0
        while True:
            try:
                name, data, vtype = winreg.EnumValue(key, i)
                type_map = {
                    winreg.REG_SZ: 'REG_SZ',
                    winreg.REG_DWORD: 'REG_DWORD',
                    winreg.REG_EXPAND_SZ: 'REG_EXPAND_SZ',
                    winreg.REG_MULTI_SZ: 'REG_MULTI_SZ',
                    winreg.REG_BINARY: 'REG_BINARY',
                    winreg.REG_QWORD: 'REG_QWORD',
                }
                results.append((name, data, type_map.get(vtype, f'type={vtype}')))
                i += 1
            except OSError:
                break
    finally:
        key.Close()
    return results


def enum_subkeys(root, path):
    """Return list of subkey names, or None if key missing."""
    try:
        key = winreg.OpenKey(root, path)
    except (FileNotFoundError, OSError):
        return None
    subs = []
    try:
        i = 0
        while True:
            try:
                subs.append(winreg.EnumKey(key, i))
                i += 1
            except OSError:
                break
    finally:
        key.Close()
    return subs


def scan_location(label, root, path, ptype):
    print(f"\n  [{label}] {path}")
    print(f"  Type: {ptype}")

    values = enum_key_values(root, path)
    if values is None:
        print(f"  -> Key does not exist")
        return False
    if values == "ACCESS_DENIED":
        print(f"  -> ACCESS DENIED (key exists but cannot read)")
        return False
    if isinstance(values, str):
        print(f"  -> {values}")
        return False

    if not values:
        print(f"  -> Key EXISTS but has no values")
    else:
        print(f"  -> Key EXISTS with {len(values)} value(s):")
        for name, data, vtype in values:
            marker = "  <<< HideRestoreDialogEnabled" if name == 'HideRestoreDialogEnabled' else ""
            print(f"       {name} = {data!r}  ({vtype}){marker}")

    # Recurse into subkeys
    subs = enum_subkeys(root, path)
    if subs:
        for sub in subs:
            sub_path = f"{path}\\{sub}"
            sub_values = enum_key_values(root, sub_path)
            if sub_values and isinstance(sub_values, list) and sub_values:
                print(f"\n  [{label}\\{sub}] {sub_path}")
                print(f"  Type: {ptype}")
                print(f"  -> {len(sub_values)} value(s):")
                for name, data, vtype in sub_values:
                    marker = "  <<< HideRestoreDialogEnabled" if name == 'HideRestoreDialogEnabled' else ""
                    print(f"       {name} = {data!r}  ({vtype}){marker}")

    return bool(values)


def search_for_value(root, start_path, target_name, results, label):
    """Recursively search for a value name under a registry subtree."""
    try:
        key = winreg.OpenKey(root, start_path)
    except (FileNotFoundError, OSError):
        return

    try:
        # Check values at this level
        i = 0
        while True:
            try:
                name, data, vtype = winreg.EnumValue(key, i)
                if name == target_name:
                    type_map = {winreg.REG_SZ: 'REG_SZ', winreg.REG_DWORD: 'REG_DWORD'}
                    results.append((label, start_path, name, data, type_map.get(vtype, vtype)))
                i += 1
            except OSError:
                break

        # Recurse into subkeys
        i = 0
        while True:
            try:
                sub = winreg.EnumKey(key, i)
                search_for_value(root, f"{start_path}\\{sub}", target_name, results, label)
                i += 1
            except OSError:
                break
    finally:
        key.Close()


def discover_custom_edge_data_dirs():
    """Discover UiPath PIP and DualEngine-registered custom Edge data dirs.
    Returns a list of (source_label, dir_path) for dirs that exist on disk."""
    local_app_data = os.environ.get('LOCALAPPDATA') or os.path.join(
        os.environ.get('USERPROFILE', ''), 'AppData', 'Local')

    candidates = []  # (source_label, path)

    # Known UiPath PIP Browser Profiles location
    uipath_path = os.path.join(local_app_data, r'UiPath\PIP Browser Profiles\Edge')
    candidates.append(('UiPath PIP Browser Profiles', uipath_path))

    # DualEngineCacheContainerTracker lists custom Edge data dirs
    tracker_keys = [
        ("HKCU",  winreg.HKEY_CURRENT_USER,  r'Software\Microsoft\Edge\DualEngineCacheContainerTracker'),
        ("HKLM",  winreg.HKEY_LOCAL_MACHINE, r'Software\Microsoft\Edge\DualEngineCacheContainerTracker'),
        ("HKCU WOW64", winreg.HKEY_CURRENT_USER,  r'Software\WOW6432Node\Microsoft\Edge\DualEngineCacheContainerTracker'),
        ("HKLM WOW64", winreg.HKEY_LOCAL_MACHINE, r'Software\WOW6432Node\Microsoft\Edge\DualEngineCacheContainerTracker'),
    ]
    for root_label, root, path in tracker_keys:
        try:
            with winreg.OpenKey(root, path) as key:
                i = 0
                while True:
                    try:
                        _, data, _ = winreg.EnumValue(key, i)
                        i += 1
                        if isinstance(data, str) and data:
                            norm = os.path.normpath(data)
                            # Walk up to the base data dir (parent of Default/Profile N)
                            parent = os.path.dirname(norm)
                            base = os.path.basename(parent)
                            if base == 'Default' or base.startswith('Profile '):
                                parent = os.path.dirname(parent)
                            candidates.append((f'DualEngine tracker ({root_label})', parent))
                    except OSError:
                        break
        except (FileNotFoundError, OSError):
            pass

    # Deduplicate, keep only dirs that exist
    seen = set()
    found = []
    for label, p in candidates:
        n = os.path.normpath(p)
        if n not in seen and os.path.isdir(n):
            found.append((label, n))
            seen.add(n)
    return found


def scan_edge_data_dir(label, data_dir):
    """Scan one Edge data dir for restore-related state: profiles,
    Preferences exit_type, session files, and Local State info_cache."""
    channel = os.path.basename(os.path.dirname(data_dir))
    print(f"\n  [{label}] {data_dir}")
    print(f"  Channel folder: {channel}")

    if not os.path.isdir(data_dir):
        print(f"  -> Directory does not exist on disk")
        return

    # Local State: per-profile exit_type in profile.info_cache
    local_state_path = os.path.join(data_dir, 'Local State')
    print(f"  Local State: {os.path.isfile(local_state_path)}")
    if os.path.isfile(local_state_path):
        try:
            with open(local_state_path, 'r', encoding='utf-8') as f:
                ls = json.load(f)
            info_cache = ls.get('profile', {}).get('info_cache', {})
            if info_cache:
                print(f"  Local State info_cache ({len(info_cache)} profile(s)):")
                for prof_name, info in info_cache.items():
                    et = info.get('exit_type', '<missing>')
                    ec = info.get('exited_cleanly', '<missing>')
                    flag = "  <<< CRASHED (restore dialog trigger)" if et == 'Crashed' else ""
                    print(f"    {prof_name}: exit_type={et!r}, exited_cleanly={ec!r}{flag}")
            else:
                print(f"  Local State info_cache: empty or missing")
        except Exception as e:
            print(f"  Local State read error: {e}")

    # Per-profile directories
    profile_dirs = []
    try:
        for name in os.listdir(data_dir):
            full = os.path.join(data_dir, name)
            if os.path.isdir(full) and (name == 'Default' or name.startswith('Profile ')):
                profile_dirs.append((name, full))
    except Exception as e:
        print(f"  Could not list {data_dir}: {e}")
        return

    print(f"  Profile dirs found: {len(profile_dirs)}")
    for name, pdir in profile_dirs:
        prefs_path = os.path.join(pdir, 'Preferences')
        print(f"\n    [{channel}/{name}]")
        print(f"      Preferences exists: {os.path.isfile(prefs_path)}")

        # Session files (legacy + current + Sessions dir)
        session_files = ['Last Tabs', 'Last Session', 'Current Tabs', 'Current Session']
        for sf in session_files:
            p = os.path.join(pdir, sf)
            exists = os.path.exists(p)
            flag = "  <<< HAS TABS TO RESTORE" if exists else ""
            print(f"      {sf}: {exists}{flag}")
        sessions_dir = os.path.join(pdir, 'Sessions')
        if os.path.isdir(sessions_dir):
            try:
                count = len(os.listdir(sessions_dir))
                print(f"      Sessions/ dir: EXISTS ({count} entries)  <<< HAS SESSION DATA")
            except Exception as e:
                print(f"      Sessions/ dir: EXISTS (read error: {e})  <<< HAS SESSION DATA")
        else:
            print(f"      Sessions/ dir: missing")

        # Preferences exit_type / exited_cleanly
        if os.path.isfile(prefs_path):
            try:
                with open(prefs_path, 'r', encoding='utf-8') as f:
                    prefs = json.load(f)
                prof = prefs.get('profile', {})
                et = prof.get('exit_type', '<missing>')
                ec = prof.get('exited_cleanly', '<missing>')
                flag = "  <<< CRASHED (restore dialog trigger)" if et == 'Crashed' else ""
                print(f"      Preferences profile.exit_type      = {et!r}{flag}")
                print(f"      Preferences profile.exited_cleanly = {ec!r}")
            except Exception as e:
                print(f"      Preferences read error: {e}")


def main():
    print("=" * 70)
    print("Edge Policy Location Scanner")
    print("Finds where edge://policy is reading from. READ-ONLY.")
    print("=" * 70)

    print("\n" + "-" * 70)
    print("PART 1: Check all known Edge policy registry locations")
    print("-" * 70)

    found_any = False
    for label, root, path, ptype in EDGE_POLICY_LOCATIONS:
        if scan_location(label, root, path, ptype):
            found_any = True

    print("\n" + "-" * 70)
    print("PART 2: Deep search for HideRestoreDialogEnabled anywhere under")
    print("        HKCU\\Software\\Microsoft  and  HKLM\\Software\\Microsoft")
    print("-" * 70)

    hits = []
    search_for_value(winreg.HKEY_CURRENT_USER,  r'Software\Microsoft', 'HideRestoreDialogEnabled', hits, "HKCU")
    search_for_value(winreg.HKEY_LOCAL_MACHINE, r'Software\Microsoft', 'HideRestoreDialogEnabled', hits, "HKLM")

    if hits:
        print(f"\n  Found HideRestoreDialogEnabled at {len(hits)} location(s):")
        for label, path, name, data, vtype in hits:
            print(f"    {label}\\{path}")
            print(f"      {name} = {data!r} ({vtype})")
    else:
        print("\n  HideRestoreDialogEnabled not found anywhere under HKCU/HKLM Software\\Microsoft.")

    print("\n" + "-" * 70)
    print("PART 3: EdgeUpdate policies (used by the Edge updater service)")
    print("-" * 70)
    for label, root, path in [
        ("HKCU", winreg.HKEY_CURRENT_USER,  r'Software\Policies\Microsoft\EdgeUpdate'),
        ("HKLM", winreg.HKEY_LOCAL_MACHINE, r'Software\Policies\Microsoft\EdgeUpdate'),
    ]:
        scan_location(label, root, path, 'mandatory')

    print("\n" + "-" * 70)
    print("PART 4: Custom Edge data directories (UiPath PIP + DualEngine)")
    print("-" * 70)
    custom_dirs = discover_custom_edge_data_dirs()
    if custom_dirs:
        print(f"\n  Found {len(custom_dirs)} custom Edge data dir(s) that exist on disk:")
        for label, p in custom_dirs:
            print(f"    - [{label}] {p}")
        print()
        for label, p in custom_dirs:
            scan_edge_data_dir(label, p)
    else:
        print("\n  No UiPath PIP or DualEngine custom Edge data directories found on disk.")
        print("  Checked: %LOCALAPPDATA%\\UiPath\\PIP Browser Profiles\\Edge")
        print("           DualEngineCacheContainerTracker (HKCU/HKLM, incl. WOW6432Node)")

    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    if hits:
        print(f"  HideRestoreDialogEnabled is set at:")
        for label, path, name, data, vtype in hits:
            print(f"    {label}\\{path} = {data}")
    else:
        print("  HideRestoreDialogEnabled is NOT set anywhere in the registry.")
        print("  If edge://policy shows it, check:")
        print("    - Is it listed as a 'Cloud' or 'Merge' source in edge://policy?")
        print("    - Is it under a different name in edge://policy?")
    if not found_any and not hits:
        print("  No Edge policy values found in any standard location.")
    if custom_dirs:
        print(f"  Custom Edge data dirs found: {len(custom_dirs)}")
        print("  If the registry policy is NOT set, the restore dialog can still")
        print("  appear from these alternate Edge profiles unless Layer 2 patches them.")

    print("\n  TIP: In edge://policy, each policy row shows its SOURCE:")
    print("    'Platform'  = from registry (one of the paths above)")
    print("    'Cloud'     = from Microsoft Edge management service")
    print("    'Merge'     = merged from multiple sources")
    print("  Look at the source column to know which tree to check.")
    print("=" * 70)


if __name__ == "__main__":
    main()