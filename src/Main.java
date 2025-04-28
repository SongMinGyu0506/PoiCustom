import bumil.custom.excel.node.ExcelHeaderNode;
import bumil.custom.excel.node.ExcelStyle;
import bumil.custom.excel.ExcelGenerator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.*;

public class Main {
    public static class UserDTO {
        private String name;
        private Integer age;
        private String city;
        private String district;

        // Getter/Setter
        public String getName() { return name; }
        public void setName(String name) { this.name = name; }

        public Integer getAge() { return age; }
        public void setAge(Integer age) { this.age = age; }

        public String getCity() { return city; }
        public void setCity(String city) { this.city = city; }

        public String getDistrict() { return district; }
        public void setDistrict(String district) { this.district = district; }
    }

    public static void main(String[] args) throws Exception {
        // 1. 헤더 트리 정의 (title=한글, fieldName=영문 DTO 필드명)
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

        // 2. 테스트 데이터 생성
        List<UserDTO> users = Arrays.asList(
                createUser("홍길동", 30, "서울", "강남구"),
                createUser("김철수", 25, "부산", "해운대구")
        );

        // 3. 엑셀 파일 생성
        Map<String, ExcelStyle> bodyStyleMap = new HashMap<>();
        ExcelGenerator.generateExcel("test.xlsx", root, users, bodyStyleMap);
        System.out.println("엑셀 생성 완료");

        // 4. 엑셀 파일 읽기 (한글 제목 → fieldName 매핑 사용)
        List<UserDTO> importedUsers = ExcelGenerator.parseExcelToDto("test.xlsx", UserDTO.class, 2, root);
        for (UserDTO user : importedUsers) {
            System.out.println(user.getName() + ", " + user.getAge() + ", " + user.getCity() + ", " + user.getDistrict());
        }
    }

    private static UserDTO createUser(String name, Integer age, String city, String district) {
        UserDTO user = new UserDTO();
        user.setName(name);
        user.setAge(age);
        user.setCity(city);
        user.setDistrict(district);
        return user;
    }
}
