# (주)풍산홀딩스 제조총괄 ERP AddOn 저장소

### 본 저장소는 (주)풍산홀딩스 제조총괄에서 사용하는 ERP의 AddOn 소스코드의 저장소입니다.
#

- __아래의 코딩 규칙을 필히 준수하십시오.__

1. 주석은 메서드나 변수 오른쪽에 한칸(스페이스) 띄우고 작성. (단, 2행 이상 코드에 대한 주석은 상단에 작성)
```C#
    if (Temp == 1) //변수에 대한 설명

    //변수 및 해당 로직에 대한 설명
    string Temp;
    Temp = Temp2;
```

2. 복합할당 사용 권장 (간결한 코드)
```C#
    Temp = Temp + 1; //지양
    Temp += 1; //추천
```
3. Visual Basic .NET 문법 혼용 금지 (아래 참조)
```C#
    Strig.Trim()
    String.Left()
    String.Right()
```
4. 클래스내 이벤트는 제일 마지막에 배치 (Raise_FormItemEvent(세부 메서드), Raise_FormMenuEvent, Raise_FormDataEvent, Raise_RightClickEvent 순)
5. 접근제한자 사용기준 준수
6. 불필요한 변수, 주석, 메서드, 클래스 사용 및 변수초기값 할당 지양
7. 빈 메서드 보존 지양
8. 숫자를 스트링으로 형변환 지양
```C#
    Convert.ToString(0) //지양
    "0" //추천
```
9. 변수선언부와 "try"문 사이에 한행 띄우기
```C#
    string Temp;

    try
    {
        ...
    }
```
10. 메서드와 메서드 사이에 한행 띄우기 (가독성을 위한 행구분은 두행이상 지양)
```C#
    private void SetTemp()
    {
        ...
    }

    private void GetTemp()
    {
        ...
    }
```
11. 변수선언은 try 문 밖에서, 변수할당은 try문 안에서 구현
```C#
    string Temp;

    try
    {
        Temp = "temp";
    }

```
12. 괄호중복 지양
```C#
    if ((Temp == 1)) //지양
```
13. 변수명, 메서드명, 클래스명 [파스칼표기법](https://docs.microsoft.com/ko-kr/dotnet/csharp/fundamentals/coding-style/coding-conventions)(.NET 권고사항) 사용 권장 (혼용 지양) (기존 VB 6.0은 제외, C#으로 신규 생성하는 화면)
```C#
    string TempString; //변수
    private void SetTemp() {} //메서드
    private class TempClass {} //클래스
```
14. and, or 사용시 &, | 지양, &&, || 지향
15. SAP B1 API 는 명시적으로 메모리 해제 (native .net Framework 객체는 Garbage Collector 동작)
```C#
SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

try
{
}
catch (Exception ex)
{    
}
finally
{
    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
}

```
16. 주석은 "//", 2개 이상의 "/" 사용 지양 (ex> "////") (단, xml주석(메서드, 클래스 상단)은 제외)
17. "{}" 생략 지양(if, for문 이후 한행만 구현시)
```C#
    //아래 방법 지양
    if (temp1 == 1)
        temp2 = 1;

    //아래 방법 추천
    if (temp1 == 1)
    {
        temp2 = 1;
    }
```
18. bool return 중복 지양
19. 배타적인 관계일 경우 단독 if 문 사용 지양(if ~ else if 사용)
20. Raise_FormItemEvent 이벤트는 각각 세부 이벤트용 메서드로 구현 필수(이벤트 필터 자동 등록 목적)


