# TUTORIAL-excel-with-python

### 실습 교재
https://wikidocs.net/135788

### Tip. 엑셀 상수를 사용하면 편함
엑셀에서는 Ctrl + 방향키 조합을 통해 
다음 내용이 있는 셀까지 점프(?)하는 기능이 있다.

이 역시 매크로로도 가능한데, 아래와 같이
아래 방향(xlDown)이나 우측 방향(xlToRight) 등의
엑셀 상수를 통해 가능하다.

```python
worksheet.Range("A1").End(constants.xlDown).Select()  # A1 기준으로 아래쪽에 있는 셀 선택

worksheet.Range("A9").End(constants.xlToRight).Select()  # A9 기준으로 우측에 있는 셀 선택
```

물론 이는 상수일 뿐이므로, xlDown 대신에 -4121이나, xlToRight 대신에 -4161을 넣어도 되긴 한다... [(참고)](https://learn.microsoft.com/en-us/office/vba/api/excel.xldirection)

그러나 숫자는 알아보기도 어렵고, 그렇다고 상수를 따로 지정할 바엔 그냥 기존에 정리된 걸 쓰는 게 훨씬 편하다.

### Tip. 엑셀 상수를 사용하기 위한 세팅

[출처](https://stackoverflow.com/questions/21465810/accessing-enumeration-constants-in-excel-com-using-python-and-win32com#answer-21534997)

- 엑셀 상수를 쓰기 위해선 다음과 같이 써야 함.
win32com.client.**gencache.EnsureDispatch**("Excel.Application")

- 그냥 win32com.client.**Dispatch**("Excel.Application")을 쓰면 AttributeError가 뜸.

- 왜냐하면 그냥 Dispatch를 쓰면 관련된 파일들이 "확정적"으로 생성되지 않기 때문임.

- 반면 EnsureDispatch를 쓰는 경우, 관련된 파일들이 생기는데, 이는 다음과 같은 경로에서 확인 가능.
`C:\\Users\\YourName\\AppData\\Local\\Temp\\gen_py\\3.11\\00020813-0000-0000-C000-000000000046x0x1x7` (이름은 바뀔 수 있음)

- 이 안에 \_\_init__.py가 있는데, 이 안에 엑셀 상수들이 나열되어 있음.
    - (예시로 이 저장소의 win32com 폴더에 있는 constants list가 해당 파일임)
