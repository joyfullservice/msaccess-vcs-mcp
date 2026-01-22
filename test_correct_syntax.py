"""Test the correct add-in calling syntax."""

import os
import win32com.client

try:
    # Create Access instance
    print("Creating Access instance...")
    app = win32com.client.Dispatch('Access.Application')
    print(f"Access Version: {app.Version}")
    
    # Open the TARGET database
    db_path = r'C:\Users\Adam Waller\Documents\GitHub\msaccess-vcs-addin\Testing\Testing.accdb'
    print(f"Opening target database: {db_path}")
    app.OpenCurrentDatabase(db_path)
    print("Database opened successfully")
    
    # Now call add-in function using correct syntax
    # Format: PathWithoutExtension.FunctionName
    addin_path = r'C:\Users\Adam Waller\AppData\Roaming\MSAccessVCS\Version Control.accda'
    addin_lib = os.path.splitext(addin_path)[0]  # Remove .accda
    
    print(f"\nAdd-in library path: {addin_lib}")
    print(f"Calling: {addin_lib}.GetVCSVersion")
    
    try:
        version = app.Run(f'{addin_lib}.GetVCSVersion')
        print(f"\n[SUCCESS] VCS Version: {version}")
    except Exception as e:
        print(f"\n[FAILED] Error: {e}")
        print(f"Error type: {type(e)}")
    
    # Clean up
    print("\nClosing database...")
    app.CloseCurrentDatabase()
    app.Quit()
    print("Done")
    
except Exception as e:
    print(f"[ERROR] {e}")
    import traceback
    traceback.print_exc()
