import os
import platform
import subprocess
import sys
import shutil

def build_executable():
    print("=" * 60)
    print("Building PDF to DOCX Converter executable...")
    print("=" * 60)
    
    # Detect operating system
    operating_system = platform.system()
    print(f"Detected operating system: {operating_system}")
    
    # Ensure PyInstaller is installed
    try:
        import PyInstaller
        print("PyInstaller is already installed.")
    except ImportError:
        print("PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # Ensure required dependencies are installed
    print("\nInstalling required dependencies...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
    
    # Clean previous build artifacts
    print("\nCleaning previous build artifacts...")
    if os.path.exists("build"):
        shutil.rmtree("build")
    if os.path.exists("dist"):
        shutil.rmtree("dist")
    if os.path.exists("PDFtoWORDConverter.spec"):
        os.remove("PDFtoWORDConverter.spec")
    
    # Determine the icon file based on the operating system
    icon_param = ""
    if operating_system == "Windows":
        if os.path.exists("icon.ico"):
            icon_param = "--icon=icon.ico"
            print("Using icon.ico for Windows executable")
        else:
            print("No icon.ico found, using default icon")
    elif operating_system == "Darwin":  # macOS
        if os.path.exists("icon.icns"):
            icon_param = "--icon=icon.icns"
            print("Using icon.icns for macOS application")
        else:
            print("No icon.icns found, using default icon")
    else:  # Linux or other
        if os.path.exists("icon.png"):
            icon_param = "--icon=icon.png"
            print("Using icon.png for Linux executable")
        else:
            print("No icon.png found, using default icon")
    
    # Build command
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--name", "PDFtoWORDConverter",
    ]
    
    # Add icon if available
    if icon_param:
        cmd.append(icon_param)
    
    # Add the main script
    cmd.append("pdf_converter.py")
    
    # Execute the PyInstaller command
    print("\nRunning PyInstaller with command:")
    print(" ".join(cmd))
    print("-" * 60)
    
    try:
        subprocess.check_call(cmd)
        
        # Print success message and location of the executable
        if operating_system == "Windows":
            exe_path = os.path.join("dist", "PDFtoWORDConverter.exe")
        elif operating_system == "Darwin":
            exe_path = os.path.join("dist", "PDFtoWORDConverter.app")
        else:
            exe_path = os.path.join("dist", "PDFtoWORDConverter")
        
        if os.path.exists(exe_path):
            print("\n" + "=" * 60)
            print(f"Build successful! Executable created at: {exe_path}")
            
            # Get the file size
            size_bytes = os.path.getsize(exe_path)
            size_mb = size_bytes / (1024 * 1024)
            
            print(f"Executable size: {size_mb:.2f} MB")
            print("=" * 60)
            
            print("\nNOTE: The executable is specific to your operating system:")
            print(f"- Windows: Creates a .exe file")
            print(f"- macOS: Creates a .app bundle")
            print(f"- Linux: Creates a binary executable")
            print("\nTo create executables for other platforms, you need to run")
            print("this build script on those platforms.")
        else:
            print("\nBuild seems to have completed, but the executable was not found at the expected location.")
    except subprocess.CalledProcessError as e:
        print(f"\nError building executable: {e}")
        sys.exit(1)

if __name__ == "__main__":
    build_executable()