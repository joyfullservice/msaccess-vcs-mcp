"""Read VCS version from database properties without executing VBA."""

import win32com.client

try:
    # Create DAO engine
    print("Creating DAO engine...")
    dao = win32com.client.Dispatch("DAO.DBEngine.120")
    
    # Open the add-in database
    addin_path = r'C:\Users\Adam Waller\AppData\Roaming\MSAccessVCS\Version Control.accda'
    print(f"Opening database: {addin_path}")
    db = dao.OpenDatabase(addin_path)
    
    # Read AppVersion property
    print("Reading AppVersion property...")
    for prop in db.Properties:
        if prop.Name == "AppVersion":
            version = prop.Value
            print(f"[SUCCESS] VCS Version from property: {version}")
            break
    else:
        print("[INFO] AppVersion property not found")
        print("Available properties:")
        for prop in db.Properties:
            try:
                print(f"  - {prop.Name}: {prop.Value}")
            except:
                print(f"  - {prop.Name}: (unable to read)")
    
    db.Close()
    print("Done")
    
except Exception as e:
    print(f"[ERROR] {e}")
    import traceback
    traceback.print_exc()
