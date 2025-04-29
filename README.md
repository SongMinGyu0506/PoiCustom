# 📄 README.md


# Excel Generator - N단계 다단계 엑셀 생성/파싱 라이브러리

Java 기반으로 구현된 **다단계(N단계) 엑셀 생성 및 파싱 유틸리티**입니다.  
Apache POI를 기반으로 하며, 복잡한 헤더 구조와 셀 병합, 스타일 지정, DTO 매핑 기능을 지원합니다.

---

## ✨ 주요 기능

- N단계 트리 형태의 엑셀 헤더 생성
- 헤더 자동 병합(Merge) 지원
- DTO 리스트 기반 바디 데이터 작성
- 각 컬럼별 스타일 지정 (폰트, 정렬, 배경색 등)
- 엑셀 파일을 읽어 DTO 리스트로 변환 (Reverse Mapping 지원)
- 한글 헤더 출력 + 내부 영문 필드 매핑 분리 가능

---

## 📦 사용 기술

- Java 8
- Apache POI 4.1.2

---

## 🚀 설치 및 빌드

```bash
git clone https://github.com/SongMinGyu0506/PoiCustom.git
cd excel-generator
```

필요 라이브러리 (Maven 기준):

```xml
<dependency>
  <groupId>org.apache.poi</groupId>
  <artifactId>poi-ooxml</artifactId>
  <version>4.1.2</version>
</dependency>
<dependency>
  <groupId>org.apache.commons</groupId>
  <artifactId>commons-compress</artifactId>
  <version>1.18</version>
</dependency>
<dependency>
  <groupId>org.apache.commons</groupId>
  <artifactId>commons-collections4</artifactId>
  <version>4.4</version>
</dependency>
<dependency>
  <groupId>org.apache.xmlbeans</groupId>
  <artifactId>xmlbeans</artifactId>
  <version>3.1.0</version>
</dependency>
```

---

## 🛠 사용 예제

### 엑셀 파일 생성

```java
ExcelHeaderNode root = new ExcelHeaderNode("사용자 정보", null, Arrays.asList(
    new ExcelHeaderNode("기본 정보", null, Arrays.asList(
        new ExcelHeaderNode("이름", "name", null, null),
        new ExcelHeaderNode("나이", "age", null, null)
    ), null),
    new ExcelHeaderNode("주소 정보", null, Arrays.asList(
        new ExcelHeaderNode("도시", "city", null, null),
        new ExcelHeaderNode("구/군", "district", null, null)
    ), null)
), null);

List<UserDTO> users = Arrays.asList(
    new UserDTO("홍길동", 30, "서울", "강남구"),
    new UserDTO("김철수", 25, "부산", "해운대구")
);

ExcelGenerator.generateExcel("test.xlsx", root, users, new HashMap<>());
```

### 엑셀 파일 파싱

```java
List<UserDTO> imported = ExcelGenerator.parseExcelToDto("test.xlsx", UserDTO.class, 2);
```

---

## 📄 클래스 구조

| 클래스명 | 역할 |
|:--|:--|
| `ExcelGenerator` | 엑셀 생성/파싱 기능 제공 |
| `ExcelHeaderNode` | 다단계 헤더 구조 정의 |
| `ExcelStyle` | 셀 스타일 지정용 모델 |

---

## 📝 라이선스

- Apache License 2.0 기반
- 자유롭게 사용 가능하며, 출처 명시 부탁드립니다.

---

# ✉️ 기여 및 문의

Pull Request와 Issue 환영합니다!  
궁금한 사항이나 제안은 언제든 연락 주세요.
