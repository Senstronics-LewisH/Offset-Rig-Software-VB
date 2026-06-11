# redeploy.py
# Simple utility to verify the compiled version inside OffsetCheck.exe and redeploy it to the Q drive network location.

import os
import sys
import shutil

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_EXE = os.path.join(PROJECT_DIR, "OffsetCheck.exe")
NETWORK_RELEASE_DIR = r"Q:\SENSTRONICS\CONTROLLED MACHINE SOFTWARE\Offset Rig Software VB"
CHANGELOG_FILE = os.path.join(PROJECT_DIR, "changelog.txt")

def main():
    print("=== Redeployment Helper for Offset Rig Software ===")
    
    # 1. Ask the user for the expected version
    expected_version = "v1.0.37"
    user_version = input(f"Enter the expected version to verify [{expected_version}]: ").strip()
    if user_version:
        expected_version = user_version
    if not expected_version.startswith("v"):
        expected_version = "v" + expected_version

    # 2. Check local exe exists
    if not os.path.exists(LOCAL_EXE):
        print(f"ERROR: Local executable '{LOCAL_EXE}' not found.")
        print("Please compile the project inside twinBASIC IDE first.")
        sys.exit(1)

    # 3. Read exe and verify version UTF-16
    print(f"Reading '{LOCAL_EXE}' to verify version '{expected_version}'...")
    version_utf16 = expected_version.encode("utf-16-le")
    
    with open(LOCAL_EXE, "rb") as f:
        exe_data = f.read()

    if version_utf16 not in exe_data:
        print("\n" + "!" * 80)
        print(f"VERIFICATION FAILED: The compiled OffsetCheck.exe does not contain the version '{expected_version}'!")
        print("It likely still contains the old version. Please follow these steps:")
        print("  1. Close the twinBASIC IDE completely.")
        print("  2. Open the twinBASIC IDE and load 'Project1.twinproj'.")
        print("  3. Compile the project ('OffsetCheck.exe') to this folder.")
        print("  4. Rerun this script.")
        print("!" * 80)
        sys.exit(1)

    print("SUCCESS: Version verified!")

    # 4. Copy to Q drive
    if not os.path.exists(NETWORK_RELEASE_DIR):
        print(f"ERROR: Network directory '{NETWORK_RELEASE_DIR}' is not accessible.")
        sys.exit(1)

    net_exe = os.path.join(NETWORK_RELEASE_DIR, "OffsetCheck.exe")
    net_changelog = os.path.join(NETWORK_RELEASE_DIR, "changelog.txt")

    try:
        shutil.copy2(LOCAL_EXE, net_exe)
        print(f"Successfully copied compiled executable to {net_exe}")
        
        if os.path.exists(CHANGELOG_FILE):
            shutil.copy2(CHANGELOG_FILE, net_changelog)
            print(f"Successfully copied changelog to {net_changelog}")
            
        print("\nRedeployment Completed successfully!")
    except Exception as e:
        print(f"ERROR: Copy failed. {str(e)}")

if __name__ == "__main__":
    main()
