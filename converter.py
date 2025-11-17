import os
import sys
import win32com.client as win32
import pythoncom
from pywintypes import com_error
import threading
import time
from pywinauto.application import Application
from pywinauto.findwindows import ElementNotFoundError

def auto_click_hwp_security_dialog():
    """
    별도의 스레드에서 한컴오피스 보안 경고창을 감시하고 '모두 허용'을 클릭합니다.
    """
    try:
        # "한컴오피스 한글" 이라는 제목의 대화상자가 나타날 때까지 최대 10초 대기
        app = Application().connect(title="한컴오피스 한글", timeout=10)
        dialog = app.window(title="한컴오피스 한글")
        
        # 대화상자가 활성화될 때까지 잠시 대기
        dialog.wait('active', timeout=5)
        
        # '모두 허용(A)' 버튼에 해당하는 Alt+A 키 입력 전송
        dialog.type_keys('%A')
        print("보안 경고창에서 '모두 허용'을 자동으로 클릭했습니다.")
        
    except ElementNotFoundError:
        # 대화상자가 나타나지 않으면 아무것도 하지 않음
        print("보안 경고창이 나타나지 않았습니다.")
    except Exception as e:
        print(f"보안 경고창 자동 클릭 중 오류 발생: {e}")

def convert_to_pdf(input_path, output_path=None):
    """
    HWP 또는 HWPX 파일을 PDF로 변환합니다.
    한컴오피스 한글이 설치되어 있어야 합니다.
    """
    hwp = None
    try:
        pythoncom.CoInitialize()

        hwp = win32.gencache.EnsureDispatch("HWPFrame.HWPObject")
        
        # 보안 모듈 적용
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        
        # 한/글 창 숨기기
        hwp.XHwpWindows.Item(0).Visible = False

        input_path = os.path.abspath(input_path)
        if output_path is None:
            output_path = os.path.splitext(input_path)[0] + ".pdf"
        else:
            output_path = os.path.abspath(output_path)
        
        if not os.path.exists(input_path):
            print(f"오류: 입력 파일 '{input_path}'을 찾을 수 없습니다.")
            return False

        # 파일 열기
        hwp.Open(input_path, "HWP", "forceopen:true")

        # PDF로 저장
        action = hwp.CreateAction("FileSaveAsPdf")
        pset = action.CreateSet()
        action.GetDefault(pset)
        pset.SetItem("FileName", output_path)
        pset.SetItem("Format", "PDF")
        action.Execute(pset)

        print(f"변환 완료: '{os.path.basename(input_path)}' -> '{os.path.basename(output_path)}'")
        return True

    except com_error as e:
        print(f"COM 오류 발생: 한컴오피스 프로그램이 설치되었는지, COM 등록이 올바른지 확인하세요.")
        print(f"상세 정보: {e}")
        return False
    except Exception as e:
        print(f"알 수 없는 오류 발생: {e}")
        return False
    finally:
        if hwp:
            hwp.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("사용법: python converter.py <input_file.hwp> [output_file.pdf]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(input_file):
        print(f"오류: 입력 파일 '{input_file}'을 찾을 수 없습니다.")
        sys.exit(1)

    convert_to_pdf(input_file, output_file)
