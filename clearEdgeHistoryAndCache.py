import json
import os
import shutil
import subprocess
import time
import winreg
import psutil

def close_edge_processes():
    """Close all Microsoft Edge processes"""
    print("Closing Microsoft Edge processes...")
    edge_processes = ['msedge.exe', 'MicrosoftEdge.exe', 'MicrosoftEdgeCP.exe', 'MicrosoftEdgeSH.exe']
    closed_count = 0
    
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] in edge_processes:
                proc.kill()
                closed_count += 1
                print(f"Closed process: {proc.info['name']}")
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    
    if closed_count > 0:
        print(f"Closed {closed_count} Edge process(es)")
        time.sleep(2)  # Wait for processes to fully close
    else:
        print("No Edge processes were running")
    
    return closed_count > 0

def find_edge_user_data_dirs():
    """Find all Edge user data directories.
    Checks: UserDataDir policy (HKLM/HKCU, incl. WOW6432Node), then default
    locations for each Edge channel (stable, Beta, Dev, Canary/SxS)."""
    local_app_data = os.environ.get('LOCALAPPDATA') or os.path.join(
        os.environ.get('USERPROFILE', ''), 'AppData', 'Local')

    candidates = []
    # UserDataDir policy can point Edge at a custom data folder
    policy_keys = [
        (winreg.HKEY_LOCAL_MACHINE, r'Software\Policies\Microsoft\Edge'),
        (winreg.HKEY_LOCAL_MACHINE, r'Software\WOW6432Node\Policies\Microsoft\Edge'),
        (winreg.HKEY_CURRENT_USER, r'Software\Policies\Microsoft\Edge'),
    ]
    for root, path in policy_keys:
        try:
            with winreg.OpenKey(root, path) as key:
                val, _ = winreg.QueryValueEx(key, 'UserDataDir')
                if val:
                    candidates.append(os.path.expandvars(val))
        except (FileNotFoundError, OSError):
            pass

    # Default per-channel locations
    candidates.extend([
        os.path.join(local_app_data, r'Microsoft\Edge\User Data'),
        os.path.join(local_app_data, r'Microsoft\Edge Beta\User Data'),
        os.path.join(local_app_data, r'Microsoft\Edge Dev\User Data'),
        os.path.join(local_app_data, r'Microsoft\Edge SxS\User Data'),
    ])

    found = []
    seen = set()
    for c in candidates:
        c_norm = os.path.normpath(c)
        if c_norm not in seen and os.path.isdir(c_norm):
            found.append(c_norm)
            seen.add(c_norm)
    return found

def hide_restore_dialog(user_profile):
    """Suppress the Edge 'Restore pages' dialog after a browser crash.
    Layers: 1) Registry policy (HideRestoreDialogEnabled), 2) Preferences JSON patch + session-file removal."""
    value_name = 'HideRestoreDialogEnabled'
    key_path = r'Software\Policies\Microsoft\Edge'

    # Layer 1: registry policy (best approach, persistent)
    try:
        key = winreg.HKEY_CURRENT_USER
        handles = []
        for part in key_path.split('\\'):
            key = winreg.CreateKey(key, part)
            handles.append(key)
        winreg.SetValueEx(key, value_name, 0, winreg.REG_DWORD, 1)
        for h in handles:
            h.Close()
        print(f"✓ Set Edge policy {value_name}=1 (HKCU) - restore dialog will be hidden")
        return 'registry'
    except PermissionError:
        try:
            subprocess.run(
                ['reg', 'add', f'HKCU\\{key_path}', '/v', value_name,
                 '/t', 'REG_DWORD', '/d', '1', '/f'],
                check=True, capture_output=True
            )
            print(f"✓ Set Edge policy {value_name}=1 (HKCU) via reg.exe")
            return 'registry'
        except subprocess.CalledProcessError as e:
            print(f"  Registry policy unavailable: {e.stderr.decode().strip()}")
    except Exception as e:
        print(f"  Registry policy unavailable: {e}")

    # Layer 2: Preferences JSON patch + delete session files (AppData, always writable)
    if not user_profile:
        return None

    # Detect Edge's actual data folder(s) - may be custom (UserDataDir policy)
    # or a non-stable channel (Beta/Dev/Canary), not just the default location.
    data_dirs = find_edge_user_data_dirs()
    if not data_dirs:
        print("  No Edge user data directory found (custom UserDataDir? check policy)")
        return None

    # Re-kill Edge: Startup Boost may have respawned background processes during cache clearing
    close_edge_processes()

    patched_count = 0
    for edge_user_data in data_dirs:
        channel = os.path.basename(os.path.dirname(edge_user_data))

        # Find all profile directories (Default, Profile 1, Profile 2, ...)
        profile_dirs = []
        for name in os.listdir(edge_user_data):
            full = os.path.join(edge_user_data, name)
            if os.path.isdir(full) and (name == 'Default' or name.startswith('Profile ')):
                profile_dirs.append(full)

        if not profile_dirs:
            print(f"  No profiles found in {channel}")
            continue

        for profile_dir in profile_dirs:
            prefs_path = os.path.join(profile_dir, 'Preferences')
            last_tabs = os.path.join(profile_dir, 'Last Tabs')
            last_session = os.path.join(profile_dir, 'Last Session')
            profile_name = os.path.basename(profile_dir)

            # Remove session files (nothing to restore = no restore dialog)
            for path, fname in [(last_tabs, 'Last Tabs'), (last_session, 'Last Session')]:
                try:
                    if os.path.exists(path):
                        os.remove(path)
                except Exception as e:
                    print(f"  Could not remove {fname} ({channel}/{profile_name}): {e}")

            # Patch Preferences with retries (file may be locked briefly after kill)
            if not os.path.isfile(prefs_path):
                continue

            patched = False
            for attempt in range(5):
                try:
                    with open(prefs_path, 'r', encoding='utf-8') as f:
                        prefs = json.load(f)
                    prefs.setdefault('profile', {})['exit_type'] = 'Normal'
                    prefs['profile']['exited_cleanly'] = True
                    with open(prefs_path, 'w', encoding='utf-8') as f:
                        json.dump(prefs, f, indent=2)
                    patched = True
                    break
                except (PermissionError, OSError):
                    if attempt < 4:
                        time.sleep(0.5)
                    else:
                        print(f"  Could not patch Preferences ({channel}/{profile_name}): file locked")

            if patched:
                print(f"✓ Patched Preferences [{channel}/{profile_name}]")
                patched_count += 1

    if patched_count > 0:
        return 'preferences'
    print("✗ Could not patch any profile's Preferences")
    return None

def clear_edge_cache_and_history():
    """Clear Microsoft Edge cache and browsing history (including IE mode)"""
    print("\n" + "="*60)
    print("Microsoft Edge Cache and History Cleaner")
    print("Clearing: Browsing History + Cached Images and Files")
    print("="*60 + "\n")
    
    # Get user profile path
    user_profile = os.environ.get('USERPROFILE')
    if not user_profile:
        print("Error: Could not find user profile directory")
        return False

    # Detect Edge's actual data folder(s) - may be custom (UserDataDir policy)
    # or a non-stable channel (Beta/Dev/Canary), not just the default location.
    data_dirs = find_edge_user_data_dirs()
    if not data_dirs:
        print("Error: No Edge user data directory found (custom UserDataDir? check policy)")
        return False
    print(f"Edge data folder(s) detected: {len(data_dirs)}")
    for d in data_dirs:
        print(f"  - {d}")

    # Build Edge data paths - ONLY History and Cache - across ALL detected data dirs and profiles
    edge_data_paths = {}
    for edge_user_data in data_dirs:
        channel = os.path.basename(os.path.dirname(edge_user_data))
        if not os.path.isdir(edge_user_data):
            continue
        for name in os.listdir(edge_user_data):
            profile_dir = os.path.join(edge_user_data, name)
            if not os.path.isdir(profile_dir):
                continue
            if not (name == 'Default' or name.startswith('Profile ')):
                continue
            tag = f"{channel}/{name}"
            edge_data_paths[f'Browsing History [{tag}]'] = os.path.join(profile_dir, 'History')
            edge_data_paths[f'History Journal [{tag}]'] = os.path.join(profile_dir, 'History-journal')
            edge_data_paths[f'Cache [{tag}]'] = os.path.join(profile_dir, 'Cache')
            edge_data_paths[f'Code Cache [{tag}]'] = os.path.join(profile_dir, 'Code Cache')
            edge_data_paths[f'GPUCache [{tag}]'] = os.path.join(profile_dir, 'GPUCache')
            edge_data_paths[f'Service Worker Cache [{tag}]'] = os.path.join(profile_dir, 'Service Worker', 'CacheStorage')
    
    # IE Mode cache paths (Edge uses these for IE mode)
    ie_mode_paths = {
        'IE Mode History': os.path.join(user_profile, r'AppData\Local\Microsoft\Windows\History'),
    }
    
    # Combine all paths
    all_paths = {**edge_data_paths, **ie_mode_paths}
    
    # Close Edge before clearing
    edge_was_running = close_edge_processes()
    
    # Clear each data location
    cleared_count = 0
    failed_count = 0
    
    print("Clearing Edge data:")
    for name, path in all_paths.items():
        try:
            if os.path.exists(path):
                if os.path.isdir(path):
                    # Remove directory and its contents
                    shutil.rmtree(path, ignore_errors=False)
                    print(f"✓ Cleared {name}: {path}")
                elif os.path.isfile(path):
                    # Remove file
                    os.remove(path)
                    print(f"✓ Cleared {name}: {path}")
                cleared_count += 1
            else:
                print(f"- {name} not found (already clear)")
        except PermissionError:
            print(f"✗ Permission denied for {name}: {path}")
            failed_count += 1
        except Exception as e:
            print(f"✗ Error clearing {name}: {e}")
            failed_count += 1
    
    # Hide restore pages dialog via Edge policy (HKCU)
    restore_result = hide_restore_dialog(user_profile)
    
    # Summary
    print("\n" + "="*60)
    print(f"Summary: {cleared_count} items cleared, {failed_count} failed")
    if restore_result == 'registry':
        print("Restore pages dialog: hidden (HideRestoreDialogEnabled=1)")
    elif restore_result == 'preferences':
        print("Restore pages dialog: hidden (Preferences patched)")
    else:
        print("Restore pages dialog: could not suppress")
    print("="*60)
    
    if failed_count > 0:
        print("\nNote: Some items could not be cleared. Make sure Edge is completely closed.")
        print("You may need to run this script as Administrator for full access.")
    
    return failed_count == 0

def main():
    try:
        success = clear_edge_cache_and_history()
        
        if success:
            print("\n✓ Edge cache and history cleared successfully!")
        else:
            print("\n⚠ Edge cache and history partially cleared (some errors occurred)")
        
    except Exception as e:
        print(f"\n✗ An error occurred: {e}")
        import traceback
        print(traceback.format_exc())

if __name__ == "__main__":
    main()
