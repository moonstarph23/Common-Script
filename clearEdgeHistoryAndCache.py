import os
import shutil
import subprocess
import time
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
    
    # Edge data paths - ONLY History and Cache
    edge_data_paths = {
        'Browsing History': os.path.join(user_profile, r'AppData\Local\Microsoft\Edge\User Data\Default\History'),
        'History Journal': os.path.join(user_profile, r'AppData\Local\Microsoft\Edge\User Data\Default\History-journal'),
        'Cache': os.path.join(user_profile, r'AppData\Local\Microsoft\Edge\User Data\Default\Cache'),
        'Code Cache': os.path.join(user_profile, r'AppData\Local\Microsoft\Edge\User Data\Default\Code Cache'),
        'GPUCache': os.path.join(user_profile, r'AppData\Local\Microsoft\Edge\User Data\Default\GPUCache'),
        'Service Worker Cache': os.path.join(user_profile, r'AppData\Local\Microsoft\Edge\User Data\Default\Service Worker\CacheStorage'),
    }
    
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
    
    # Summary
    print("\n" + "="*60)
    print(f"Summary: {cleared_count} items cleared, {failed_count} failed")
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
