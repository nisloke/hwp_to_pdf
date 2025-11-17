import sys
from win32com.client import makepy
import os

def generate_hwp_wrapper():
    """
    Programmatically runs makepy for the HWP COM object.
    """
    print("Attempting to generate COM wrapper for 'HwpObject 1.0 Type Library'...")
    try:
        # The makepy utility can be invoked with the description of the type library.
        # We use the '-i' flag to invite user input if multiple libraries match.
        sys.argv = ["makepy", "-i", "HwpObject 1.0 Type Library"]
        makepy.main()
        print("\n[Success] COM wrapper generation process completed.")
        print("If no errors were shown above, the wrapper was likely created or updated successfully.")
    except Exception as e:
        print(f"\n[Failure] An error occurred during the wrapper generation process: {e}")
        print("This might happen if the HWP Type Library is not found in the system registry.")

if __name__ == "__main__":
    print("This script will attempt to regenerate the pywin32 COM cache for HWP.")
    
    # gen_py 폴더 경로 확인 및 삭제 안내
    gen_py_path = os.path.join(os.environ['LOCALAPPDATA'], 'Temp', 'gen_py')
    if os.path.exists(gen_py_path):
        print(f"Found 'gen_py' cache folder at: {gen_py_path}")
        print("It is highly recommended to delete this folder for a clean regeneration.")
        # 사용자에게 직접 삭제하도록 안내하는 것이 안전합니다.
        # shutil.rmtree(gen_py_path)
    
    print("\nPlease ensure you have deleted the 'gen_py' folder before proceeding if it exists.")
    
    generate_hwp_wrapper()
