import os
import sys
import subprocess

def main():
    print("=== Step 1: Building docgit.exe using PyInstaller ===")
    try:
        subprocess.run([sys.executable, "-m", "PyInstaller", "docgit.spec"], check=True)
    except Exception as e:
        print("Failed to run PyInstaller for docgit. Ensure you have installed it: pip install pyinstaller")
        print(e)
        return
        
    print("\n=== Step 2: Building docgit_server.exe using PyInstaller ===")
    try:
        subprocess.run([sys.executable, "-m", "PyInstaller", "--noconsole", "--onefile", "--add-data", "static;static", "docgit_server.py"], check=True)
    except Exception as e:
        print("Failed to run PyInstaller for docgit_server.")
        print(e)
        return
        
    if not os.path.exists("dist/docgit_server.exe"):
        print("\nError: docgit_server.exe was not created! Please fix the errors in Step 2 before building the installer.")
        return
        
    print("\n=== Step 3: Compiling Inno Setup Installer ===")
    iscc_path = r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    if not os.path.exists(iscc_path):
        iscc_path = r"C:\Program Files\Inno Setup 6\ISCC.exe"
        
    if not os.path.exists(iscc_path):
        print("\nError: Inno Setup compiler (ISCC.exe) not found.")
        print("Please download and install Inno Setup 6 from: https://jrsoftware.org/isinfo.php")
        return
        
    try:
        subprocess.run([iscc_path, "docgit.iss"], check=True)
        print("\nSUCCESS! Installer has been generated in the Output/ folder.")
    except Exception as e:
        print("\nFailed to compile Inno Setup script.")
        print(e)

if __name__ == "__main__":
    main()
