"""Test using just the filename vs full path."""

import os
import win32com.client

try:
    # Create Access instance
    print("Creating Access instance...")
    app = win32com.client.Dispatch('Access.Application')
    
    # Open the TARGET database
    db_path = r'C:\Users\Adam Waller\Documents\GitHub\msaccess-vcs-addin\Testing\Testing.accdb'
    print(f"Opening target database: {db_path}")
    app.OpenCurrentDatabase(db_path)
    print("Database opened successfully\n")
    
    addin_path = r'C:\Users\Adam Waller\AppData\Roaming\MSAccessVCS\Version Control.accda'
    
    # Test different formats
    tests = [
        ("Full path without extension", os.path.splitext(addin_path)[0]),
        ("Just filename without extension", "Version Control"),
        ("With quotes - full path", f'"{os.path.splitext(addin_path)[0]}"'),
        ("With quotes - filename only", '"Version Control"'),
    ]
    
    for name, lib_path in tests:
        print(f"TEST: {name}")
        print(f"  Calling: {lib_path}.Preload")
        try:
            result = app.Run(f'{lib_path}.Preload')
            print(f"  [SUCCESS] Result: {result}\n")
            break  # If one works, use it
        except Exception as e:
            print(f"  [FAILED] Error code: {e.args[0] if e.args else 'unknown'}\n")
    
    # Clean up
    print("Closing database...")
    app.CloseCurrentDatabase()
    app.Quit()
    print("Done")
    
except Exception as e:
    print(f"[ERROR] {e}")
    import traceback
    traceback.print_exc()
