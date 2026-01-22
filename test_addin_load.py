"""Test loading the VCS add-in with a database open."""

import os
import win32com.client

try:
    # Create Access instance
    print("Creating Access instance...")
    app = win32com.client.Dispatch('Access.Application')
    app.Visible = False
    print(f"Access Version: {app.Version}")
    
    # Open a database
    db_path = r'C:\Users\Adam Waller\Documents\GitHub\msaccess-vcs-addin\Testing\Testing.accdb'
    print(f"Opening database: {db_path}")
    app.OpenCurrentDatabase(db_path)
    print("Database opened successfully")
    
    # Try to load add-in
    addin_path = r'C:\Users\Adam Waller\AppData\Roaming\MSAccessVCS\Version Control.accda'
    addin_lib = os.path.splitext(addin_path)[0]
    
    print(f"Loading add-in from: {addin_path}")
    print(f"Calling: {addin_lib}.Preload")
    
    try:
        result = app.Run(f'"{addin_lib}.Preload"')
        print(f"Preload result: {result}")
        print("[SUCCESS] Add-in loaded!")
        
        # Try to get version
        print("Getting VCS version...")
        version = app.Run(f'"{addin_lib}.GetVCSVersion"')
        print(f"[SUCCESS] VCS Version: {version}")
        
    except Exception as e:
        print(f"[FAILED] Error loading add-in: {e}")
    
    # Clean up
    print("Closing database...")
    app.CloseCurrentDatabase()
    app.Quit()
    print("Done")
    
except Exception as e:
    print(f"[ERROR] {e}")
    import traceback
    traceback.print_exc()
