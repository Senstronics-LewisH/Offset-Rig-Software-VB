# release.py
# Developer automation script for packaging, tagging, building, and deploying Offset Rig Software releases.

import os
import re
import sys
import shutil
import subprocess
from datetime import datetime

# Configuration
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
LOGFILE_BAS = os.path.join(PROJECT_DIR, "LogFile.bas")
SYNC_SCRIPT = os.path.join(PROJECT_DIR, "sync_all_twinprojs.py")
LOCAL_EXE = os.path.join(PROJECT_DIR, "OffsetCheck.exe")
NETWORK_RELEASE_DIR = r"Q:\SENSTRONICS\CONTROLLED MACHINE SOFTWARE\Offset Rig Software VB"
CHANGELOG_FILE = os.path.join(PROJECT_DIR, "changelog.txt")

def run_git(args):
    """Helper to run git commands and return output."""
    try:
        res = subprocess.run(["git"] + args, capture_output=True, text=True, check=True, cwd=PROJECT_DIR)
        return res.stdout.strip()
    except subprocess.CalledProcessError as e:
        print(f"Git command failed: git {' '.join(args)}")
        print(e.stderr)
        return None

def get_last_tag():
    """Gets the most recent git tag."""
    tags = run_git(["tag", "--list"])
    if not tags:
        return None
    # Get last tag sorted by semantic version or creation time
    last_tag = run_git(["describe", "--tags", "--abbrev=0"])
    return last_tag

def get_commits_since(tag):
    """Gets all commits since a given tag."""
    if tag:
        range_spec = f"{tag}..HEAD"
        commits = run_git(["log", range_spec, "--pretty=format:* %s (%h)"])
    else:
        commits = run_git(["log", "--pretty=format:* %s (%h)"])
    
    if not commits:
        return []
    return [line.strip() for line in commits.splitlines() if line.strip()]

def increment_version(tag):
    """Suggests an incremented version based on the semantic tag (e.g. v1.0.34 -> v1.0.35)."""
    if not tag:
        return "v1.0.0"
    
    match = re.search(r"v?(\d+)\.(\d+)\.(\d+)", tag)
    if match:
        major, minor, patch = map(int, match.groups())
        return f"v{major}.{minor}.{patch + 1}"
    
    return tag + "-next"

def update_logfile_version(new_version):
    """Updates the CompileVersion constant inside LogFile.bas."""
    if not os.path.exists(LOGFILE_BAS):
        print(f"ERROR: {LOGFILE_BAS} not found.")
        sys.exit(1)
        
    with open(LOGFILE_BAS, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()
        
    pattern = r'(Public\s+Const\s+CompileVersion\s+As\s+String\s*=\s*")[^"]*(")'
    replacement = rf'\1{new_version}\2'
    
    new_content, count = re.subn(pattern, replacement, content)
    if count == 0:
        print("WARNING: Could not find 'Public Const CompileVersion As String' in LogFile.bas.")
        return False
        
    with open(LOGFILE_BAS, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(new_content)
        
    print(f"Updated CompileVersion to '{new_version}' in LogFile.bas")
    return True

def main():
    print("=== Offset Rig Software Release Automation ===")
    
    # 1. Verify working directory is clean
    git_status = run_git(["status", "--porcelain"])
    if git_status:
        print("WARNING: You have uncommitted changes in your repository:")
        print(git_status)
        confirm = input("Do you want to continue anyway? (y/N): ").strip().lower()
        if confirm != 'y':
            print("Aborted.")
            sys.exit(0)
            
    # 2. Get Last Tag and Commits
    last_tag = get_last_tag()
    print(f"Last Git Tag detected: {last_tag if last_tag else 'None'}")
    
    commits = get_commits_since(last_tag)
    if not commits:
        print("No new commits since last tag.")
    else:
        print(f"\nNew commits since last tag ({len(commits)}):")
        for commit in commits[:10]:
            print(f"  {commit}")
        if len(commits) > 10:
            print(f"  ... and {len(commits) - 10} more")
            
    # 3. Version Selection
    suggested = increment_version(last_tag)
    user_version = input(f"\nEnter new version version [{suggested}]: ").strip()
    new_version = user_version if user_version else suggested
    if not new_version.startswith("v"):
        new_version = "v" + new_version
        
    # 4. Generate Changelog
    changelog_entries = [
        f"Version {new_version} - {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "=" * 40
    ]
    if commits:
        changelog_entries.extend(commits)
    else:
        changelog_entries.append("* No commit messages recorded (manual release).")
    changelog_entries.append("\n")
    
    changelog_text = "\n".join(changelog_entries)
    
    # Write changelog.txt
    if os.path.exists(CHANGELOG_FILE):
        with open(CHANGELOG_FILE, "r", encoding="utf-8") as f:
            old_log = f.read()
        changelog_text = changelog_text + old_log
        
    with open(CHANGELOG_FILE, "w", encoding="utf-8") as f:
        f.write(changelog_text)
    print(f"Generated/Updated changelog at: {CHANGELOG_FILE}")

    # 5. Inject Version into code and Run Sync Script
    update_logfile_version(new_version)
    
    print("\nRunning project sync script to pack changes into twinproj...")
    try:
        subprocess.run([sys.executable, SYNC_SCRIPT], check=True)
    except subprocess.CalledProcessError as e:
        print("ERROR: Sync script failed.")
        sys.exit(1)
        
    # 6. Instruct Manual twinBASIC Compile
    print("\n" + "*" * 70)
    print("COMPILATION STEP:")
    print("Please open the twinBASIC IDE and compile the project:")
    print(f"  - Output Executable name MUST be: 'OffsetCheck.exe'")
    print(f"  - Output Executable directory MUST be: '{PROJECT_DIR}'")
    print("*" * 70)
    
    input("\nOnce compilation is completed and 'OffsetCheck.exe' is created, press Enter to continue...")
    
    if not os.path.exists(LOCAL_EXE):
        print(f"ERROR: Compiled executable '{LOCAL_EXE}' could not be found.")
        print("Please compile the project to the correct folder and rerun this script.")
        sys.exit(1)
        
    # Verify the version in the compiled binary to ensure the IDE was reloaded
    print("Verifying version string in compiled executable...")
    version_utf16 = new_version.encode("utf-16-le")
    
    while True:
        with open(LOCAL_EXE, "rb") as f:
            exe_data = f.read()
            
        if version_utf16 in exe_data:
            print(f"SUCCESS: Verified version '{new_version}' is compiled into the executable.")
            break
        else:
            print("\n" + "!" * 80)
            print(f"ERROR: The compiled executable does not contain the version string '{new_version}'!")
            print("This usually happens because the twinBASIC IDE was already open and compiled the old version from memory.")
            print("To fix this:")
            print("  1. Close the twinBASIC IDE (or close the active project/workspace).")
            print("  2. Open the project again in twinBASIC (so it loads the updated 'Project1.twinproj' from disk).")
            print("  3. Build the project again (compile to 'OffsetCheck.exe').")
            print("!" * 80)
            
            retry = input("\nHave you recompiled the project? Press Enter to retry verification (or type 'q' to quit): ").strip().lower()
            if retry == 'q':
                sys.exit(0)
                
            if not os.path.exists(LOCAL_EXE):
                print(f"ERROR: Compiled executable '{LOCAL_EXE}' could not be found.")
                continue

    # 7. Commit changes, tag, and push
    print("\nCommitting changes and creating Git tag...")
    run_git(["add", "."])
    run_git(["commit", "-m", f"Bump version to {new_version} and update changelog"])
    run_git(["tag", "-a", new_version, "-m", f"Release version {new_version}"])
    
    print("Pushing tags and commits to remote...")
    run_git(["push", "origin", "main"])
    run_git(["push", "origin", new_version])
    
    # 8. Deploy to Network Share
    print(f"\nDeploying to Network Release directory: {NETWORK_RELEASE_DIR}...")
    if not os.path.exists(NETWORK_RELEASE_DIR):
        print(f"WARNING: Network directory '{NETWORK_RELEASE_DIR}' is not accessible.")
        print("Files were NOT copied. Please manually copy 'OffsetCheck.exe' and 'changelog.txt' to the Q drive once it is available.")
    else:
        net_exe = os.path.join(NETWORK_RELEASE_DIR, "OffsetCheck.exe")
        net_changelog = os.path.join(NETWORK_RELEASE_DIR, "changelog.txt")
        
        try:
            # Copy executable
            shutil.copy2(LOCAL_EXE, net_exe)
            print(f"  Copied executable to {net_exe}")
            
            # Copy changelog
            shutil.copy2(CHANGELOG_FILE, net_changelog)
            print(f"  Copied changelog to {net_changelog}")
            print("\nDeployment Completed Successfully!")
        except Exception as e:
            print(f"ERROR: Could not copy files to network share. {str(e)}")
            
if __name__ == "__main__":
    main()
