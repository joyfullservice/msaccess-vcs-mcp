"""Test calling Preload first, then GetVCSVersion."""

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
    print("Database opened successfully\n")
    
    # Get add-in library path
    addin_path = r'C:\Users\Adam Waller\AppData\Roaming\MSAccessVCS\Version Control.accda'
    addin_lib = os.path.splitext(addin_path)[0]
    
    print(f"Add-in library path: {addin_lib}\n")
    
    # Test 1: Call Preload
    print("TEST 1: Calling Preload...")
    try:
        result = app.Run(f'{addin_lib}.Preload')
        print(f"[SUCCESS] Preload returned: {result}")
    except Exception as e:
        print(f"[FAILED] Preload error: {e}\n")
    
    # Test 2: Call GetVCSVersion
    print("\nTEST 2: Calling GetVCSVersion...")
    try:
        version = app.Run(f'{addin_lib}.GetVCSVersion')
        print(f"[SUCCESS] VCS Version: {version}")
    except Exception as e:
        print(f"[FAILED] GetVCSVersion error: {e}")
    
    # Test 3: Try HandleRibbonCommand (the one from the example)
    print("\nTEST 3: Trying HandleRibbonCommand with invalid param...")
    try:
        result = app.Run(f'{addin_lib}.HandleRibbonCommand', "test")
        print(f"Result: {result}")
    except Exception as e:
        print(f"Expected error (invalid param): {e}")
    
    # Clean up
    print("\nClosing database...")
    app.CloseCurrentDatabase()
    app.Quit()
    print("Done")
    
except Exception as e:
    print(f"[ERROR] {e}")
    import traceback
    traceback.print_exc()
