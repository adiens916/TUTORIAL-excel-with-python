# Range 메서드 & 속성들 중 쓸 만한 것들


## Method

### Range 반환하는 메서드
- [Find](https://learn.microsoft.com/en-us/office/vba/api/excel.range.find) => '찾기' 기능
    - 전체를 한꺼번에 반환하는 기능은 없는 듯.
    - FindNext, FindPrevious로 하나씩 찾아야 하는 듯함.
- [FindNext](https://learn.microsoft.com/en-us/office/vba/api/excel.range.findnext) 
    - 검색이 끝나면 첫 셀로 돌아가는데, 어떻게 이를 끝내는지 예시가 있어서 좋음. (첫 셀 주소 참조)

### 잘 모르겠음
- [Consolidate](https://learn.microsoft.com/en-us/office/vba/api/excel.range.consolidate)
- [DataSeries](https://learn.microsoft.com/en-us/office/vba/api/excel.range.dataseries)


## Property
### Range 반환하는 속성
- [Offset](https://learn.microsoft.com/en-us/office/vba/api/excel.range.offset) => 오프셋
- [Item](https://learn.microsoft.com/en-us/office/vba/api/excel.range.item) => 오프셋
    - Offset과 Item의 차이 [참고](https://blog.naver.com/rosa0189/60145630004):
        - Offset은 (0, 0)이 기준 위치고, Item은 (1, 1)이 기준 위치임.
        - Item은 VBA에서 축약 표현을 쓸 수 있음.
    - 개인적으로는 Offset이 명시적이고 덜 헷갈려서 좋은 것 같음.
        
- [Next](https://learn.microsoft.com/en-us/office/vba/api/excel.range.next) => 다음 셀
- [Previous](https://learn.microsoft.com/en-us/office/vba/api/excel.range.previous) => 이전 셀

### 변수 반환하는 속성
- [Address](https://learn.microsoft.com/en-us/office/vba/api/excel.range.address) => 셀 주소 반환
- [Interior](https://learn.microsoft.com/en-us/office/vba/api/excel.range.interior) => 셀 색 변경
- [Left](https://learn.microsoft.com/en-us/office/vba/api/excel.range.left) => A열까지의 거리
- [Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2) => 형식 미적용 값

### 잘 모르겠음
- [AddressLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.range.addresslocal)
- [Areas](https://learn.microsoft.com/en-us/office/vba/api/excel.range.areas)
- [Cells](https://learn.microsoft.com/en-us/office/vba/api/excel.range.cells)
- [CurrentArray](https://learn.microsoft.com/en-us/office/vba/api/excel.range.currentarray)
- [CurrentRegion](https://learn.microsoft.com/en-us/office/vba/api/excel.range.currentregion)
- [ListObject](https://learn.microsoft.com/en-us/office/vba/api/excel.range.listobject)
