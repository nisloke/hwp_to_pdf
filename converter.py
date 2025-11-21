import os
import sys
import win32com.client as win32
import pythoncom
from pywintypes import com_error
import threading
import time
from pywinauto import Desktop

def monitor_security_dialog():
    """
    별도의 스레드에서 한컴오피스 보안 경고창을 감시하고 '모두 허용(A)'을 클릭합니다.
    """
    start_time = time.time()
    while time.time() - start_time < 10:  # 감시 시간을 10초로 연장
        try:
            desktop = Desktop(backend="win32")
            
            # 모든 활성 창을 순회하며 보안 경고창 찾기
            target_dialog = None
            for win in desktop.windows():
                try:
                    title = win.window_text()
                    # 제목에 '한컴오피스'나 '한글'이 있고, 내용 추정을 위해 접근
                    if "한컴오피스" in title or "한글" in title:
                        # 창 내부 텍스트나 버튼을 확인하여 보안 경고창인지 확신하면 좋음
                        # 여기서는 제목 조건이 맞으면 시도
                        target_dialog = win
                        break
                except Exception:
                    continue
            
            if target_dialog:
                # 창을 찾았으면 활성화하고 키 전송
                try:
                    if not target_dialog.is_active():
                        target_dialog.set_focus()
                    
                    # '모두 허용' 단축키 Alt+A 전송
                    target_dialog.type_keys('%A', with_spaces=True)
                    print(f"보안 경고창('{target_dialog.window_text()}') 감지 및 '모두 허용(Alt+A)' 전송 완료")
                    break
                except Exception as e:
                    print(f"창 제어 실패 (재시도 중): {e}")

        except Exception:
            pass
        
        time.sleep(0.5)

def convert_to_pdf(input_path, output_path=None):
    """
    HWP 또는 HWPX 파일을 PDF로 변환합니다.
    한컴오피스 한글이 설치되어 있어야 합니다.
    """
    hwp = None
    try:
        pythoncom.CoInitialize()

        hwp = win32.gencache.EnsureDispatch("HWPFrame.HWPObject")
        
        # 보안 모듈 적용 (RegisterModule로 해결되면 가장 좋음)
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

        # 보안 경고창 감시 스레드 시작 (Open 호출 직전)
        security_thread = threading.Thread(target=monitor_security_dialog, daemon=True)
        security_thread.start()

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
