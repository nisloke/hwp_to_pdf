import win32com.client
import pythoncom
from pywintypes import com_error
import os

def test_hwp_conversion():
    """
    HWP COM 객체를 사용하여 HWP 파일을 PDF로 변환하는 전체 과정을 테스트합니다.
    """
    hwp = None
    try:
        # COM 라이브러리 초기화
        pythoncom.CoInitialize()
        print("COM 라이브러리를 초기화했습니다.")

        # HWP 파일 경로 설정
        script_path = os.path.dirname(os.path.abspath(__file__))
        hwp_file = os.path.join(script_path, "TestBTL.hwp")
        pdf_file = os.path.join(script_path, "TestBTL.pdf")

        print(f"입력 파일: {hwp_file}")
        print(f"출력 파일: {pdf_file}")

        if not os.path.exists(hwp_file):
            print(f"\n[실패] HWP 파일이 존재하지 않습니다: {hwp_file}")
            return

        # HWP COM 객체 생성 시도
        print("\nHWPFrame.HWPObject 객체 생성을 시도합니다...")
        hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HWPObject")
        
        if hwp:
            print("[성공] HWP COM 객체를 성공적으로 생성했습니다.")
            
            # 한/글 창을 보이게 설정 (디버깅용)
            hwp.XHwpWindows.Item(0).Visible = True
            print("한/글 창을 화면에 표시합니다.")

            # 파일 열기
            print(f"'{os.path.basename(hwp_file)}' 파일 열기를 시도합니다...")
            hwp.Open(hwp_file, "HWP", "forceopen:true")
            print("[성공] 파일을 열었습니다.")

            # PDF로 저장 (CreateAction 사용)
            print(f"'{os.path.basename(pdf_file)}' PDF 파일로 저장합니다...")
            action = hwp.CreateAction("FileSaveAsPdf")
            pset = action.CreateSet()
            action.GetDefault(pset) # 기본값 로드
            pset.SetItem("FileName", pdf_file)
            pset.SetItem("Format", "PDF")
            action.Execute(pset)
            print("[성공] PDF 파일로 저장했습니다.")
            
            # 생성된 객체를 정상적으로 종료
            hwp.Quit()
            print("\nHWP 객체를 종료했습니다.")
        else:
            print("\n[실패] 객체는 생성되었으나 유효하지 않습니다.")

    except com_error as e:
        print(f"\n[실패] COM 오류가 발생했습니다.")
        print("이 오류는 한컴오피스 한글이 설치되지 않았거나,")
        print("프로그램의 COM 등록이 손상되었을 때 발생할 수 있습니다.")
        print(f" - 오류 정보: {e}")

    except Exception as e:
        print(f"\n[실패] 예기치 않은 오류가 발생했습니다: {e}")

    finally:
        # COM 라이브러리 해제
        if 'pythoncom' in locals() or 'pythoncom' in globals():
            pythoncom.CoUninitialize()
            print("COM 라이브러리를 해제했습니다.")

if __name__ == "__main__":
    test_hwp_conversion()

