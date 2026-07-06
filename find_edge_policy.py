#!/usr/bin/env python3
"""Find where Edge policies are actually saved in the registry.

Edge reads policies from FOUR registry trees (mandatory + recommended,
each in HKLM and HKCU), plus their WOW6432Node mirrors. This script
scans all of them and prints every value found, with the full path,
so you can see exactly where edge://policy is pulling from.

READ-ONLY - changes nothing.
"""

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

    print("\n  TIP: In edge://policy, each policy row shows its SOURCE:")
    print("    'Platform'  = from registry (one of the paths above)")
    print("    'Cloud'     = from Microsoft Edge management service")
    print("    'Merge'     = merged from multiple sources")
    print("  Look at the source column to know which tree to check.")
    print("=" * 70)


if __name__ == "__main__":
    main()