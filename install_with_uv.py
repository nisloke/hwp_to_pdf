import subprocess
import sys
import os

def install_packages_with_uv():
    """
    Installs required packages using uv.
    """
    packages = [
        "pyhwp",
        "pywin32",
        "customtkinter",
        "tkinterdnd2", # tkinterdnd2-universal 대신 tkinterdnd2 사용
        "pyinstaller"
    ]

    print(f"Installing packages with uv: {', '.join(packages)}")
    
    # uv 실행 파일 경로
    uv_executable = os.path.join(sys.prefix, "Scripts", "uv.exe")
    if not os.path.exists(uv_executable):
        print(f"Error: uv executable not found at {uv_executable}")
        print("Please ensure uv is installed in your virtual environment.")
        sys.exit(1)

    try:
        # uv pip install 명령 실행
        command = [uv_executable, "pip", "install", "--python", sys.executable] + packages
        subprocess.run(command, check=True, capture_output=True, text=True)
        print("\n[Success] All required packages installed successfully with uv.")
    except subprocess.CalledProcessError as e:
        print(f"\n[Failure] Error installing packages with uv:")
        print(f"Stdout: {e.stdout}")
        print(f"Stderr: {e.stderr}")
        sys.exit(1)
    except Exception as e:
        print(f"\n[Failure] An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    install_packages_with_uv()
