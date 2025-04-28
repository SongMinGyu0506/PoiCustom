# Excel Generator API 문서

## 📦 패키지 구조

| 패키지 | 설명 |
|:--|:--|
| `ExcelGenerator` | 엑셀 파일 생성 및 파싱 기능 제공 |
| `ExcelHeaderNode` | 엑셀 헤더 구조 정의 (트리 형태 지원) |
| `ExcelStyle` | 셀 스타일 지정용 모델 클래스 |

---

## 📚 주요 클래스 및 메소드

### ExcelHeaderNode

트리 형태로 다단계(N단계) 엑셀 헤더를 구성하는 클래스입니다.

- `title` : 엑셀에 표시할 컬럼 이름 (예: `"이름"`, `"나이"`)
- `fieldName` : DTO 매핑용 필드명 (예: `"name"`, `"age"`)
- `children` : 하위 헤더 노드 리스트
- `style` : 셀 스타일 정보

### ExcelStyle

엑셀 셀에 적용할 스타일을 정의하는 클래스입니다.

- `fontName` : 폰트명
- `fontSize` : 폰트 크기
- `bold` : 볼드체 여부
- `backgroundColor` : 배경 색상
- `alignment` : 가로 정렬
- `verticalAlignment` : 세로 정렬

### ExcelGenerator

엑셀 생성 및 읽기를 담당하는 유틸리티 클래스입니다.

#### generateExcel

```java
public static void generateExcel(
    String filename,
    ExcelHeaderNode headerRoot,
    List<?> dataList,
    Map<String, ExcelStyle> bodyStyleMap
) throws Exception
```

- `filename` : 저장할 엑셀 파일명 (예: `"sample.xlsx"`)
- `headerRoot` : 헤더 트리 최상위 노드
- `dataList` : 바디 영역에 매핑할 DTO 리스트
- `bodyStyleMap` : 컬럼명 기준 바디 셀 스타일 매핑 (선택)

#### parseExcelToDto

```java
public static <T> List<T> parseExcelToDto(
    String filename,
    Class<T> dtoClass,
    int headerEndRow
) throws Exception
```

- `filename` : 읽을 엑셀 파일명
- `dtoClass` : 매핑할 DTO 클래스
- `headerEndRow` : 헤더가 끝나는 라인 번호 (0부터 시작)

---

## 🧩 사용 흐름

1. `ExcelHeaderNode`로 헤더 트리 정의
2. DTO 리스트 준비
3. `generateExcel()`로 엑셀 생성
4. `parseExcelToDto()`로 엑셀 파일 읽어 DTO 리스트 변환

---

## 💬 주의사항

- 엑셀에 표시되는 `title`은 자유롭게 작성 가능(한글 가능)
- 내부 매핑용 `fieldName`은 DTO 필드명과 정확히 일치해야 함
- POI 및 관련 라이브러리 의존성 필요