"""Test calling VCS add-in functions by opening it directly."""

import os
import win32com.client

try:
    # Create Access instance
    print("Creating Access instance...")
    app = win32com.client.Dispatch('Access.Application')
    app.Visible = False
    print(f"Access Version: {app.Version}")
    
    # Open the add-in database directly
    addin_path = r'C:\Users\Adam Waller\AppData\Roaming\MSAccessVCS\Version Control.accda'
    print(f"Opening add-in database: {addin_path}")
    app.OpenCurrentDatabase(addin_path)
    print("Add-in database opened successfully")
    
    # Now try to call the functions directly
    try:
        print("Calling Preload...")
        result = app.Run("Preload")
        print(f"Preload result: {result}")
        
        print("Calling GetVCSVersion...")
        version = app.Run("GetVCSVersion")
        print(f"[SUCCESS] VCS Version: {version}")
        
    except Exception as e:
        print(f"[FAILED] Error calling function: {e}")
    
    # Clean up
    print("Closing database...")
    app.CloseCurrentDatabase()
    app.Quit()
    print("Done")
    
except Exception as e:
    print(f"[ERROR] {e}")
    import traceback
    traceback.print_exc()
