# TUTORIAL-excel-with-python

[VBA 엑셀 상수를 쓰는 방법]

[출처](https://stackoverflow.com/questions/21465810/accessing-enumeration-constants-in-excel-com-using-python-and-win32com#answer-21534997)

- 엑셀 상수를 쓰기 위해선 다음과 같이 써야 함.
win32com.client.**gencache.EnsureDispatch**("Excel.Application")

- 그냥 win32com.client.**Dispatch**("Excel.Application")을 쓰면 AttributeError가 뜸.

- 왜냐하면 그냥 Dispatch를 쓰면 관련된 파일들이 "확정적"으로 생성되지 않기 때문임.

- 반면 EnsureDispatch를 쓰는 경우, 관련된 파일들이 생기는데, 이는 다음과 같은 경로에서 확인 가능.
`C:\\Users\\YourName\\AppData\\Local\\Temp\\gen_py\\3.11\\00020813-0000-0000-C000-000000000046x0x1x7` (이름은 바뀔 수 있음)

- 이 안에 \_\_init__.py가 있는데, 이 안에 엑셀 상수들이 나열되어 있음.
    - (예시로 이 저장소의 win32com 폴더에 있는 constants list가 해당 파일임)
