# Image Excel Organizer

이미지 파일을 파일명 규칙에 따라 자동 정렬하고 Excel 표에 배치하는 Windows용 GUI 프로그램입니다.

## 주요 기능
- TIFF/TIF, PNG, JPEG/JPG, BMP, GIF, WebP 등 다양한 이미지 입력
- 파일명 기준 자연 정렬 및 행 그룹화
- 이미지 정렬표 + 파일명 검증표 생성
- 이미지 셀 크기 사용자 설정 (기본 열 너비 17 / 행 높이 80)
- 셀 크기에 맞춰 이미지 표시 크기 자동 조정
- 원본 픽셀을 리사이즈하지 않는 화질 유지 방식
- 내보내기 전 정렬 미리보기
- 같은 그룹 내 수동 순서 조정
- 설정값 저장/복원
- 중복 파일명 및 선택형 순번 누락 검사
- 진행률 표시 및 취소
- 생성 완료 후 Excel/저장 폴더 바로 열기
- tkinterdnd2 설치 시 드래그앤드롭 지원

## 실행 환경
- Windows
- Python 3.10 이상

## 설치
```powershell
python -m pip install -r requirements.txt
```

## 실행
```powershell
python excel_image_organizer_final_ux.py
```

## 테스트
```powershell
python -m unittest test_excel_image_organizer_ux.py
```

## 다운로드
GitHub 화면에서 **Code → Download ZIP**을 선택하면 이 저장소의 전체 파일을 ZIP으로 받을 수 있습니다.

## 참고
드래그앤드롭은 `tkinterdnd2`가 설치된 환경에서 활성화됩니다. 설치되지 않은 경우에도 기존 파일 선택 방식으로 정상 실행됩니다.
