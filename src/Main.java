import bumil.custom.excel.ExcelGenerator;
import bumil.custom.excel.node.ExcelHeaderNode;

import java.io.*;
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
        // 1. 헤더 트리 구성
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

        // 2. 테스트 데이터 준비
        List<UserDTO> users = Arrays.asList(
                createUser("홍길동", 30, "서울", "강남구"),
                createUser("김철수", 25, "부산", "해운대구")
        );

        // 3. 엑셀 파일 생성
        try (OutputStream fos = new FileOutputStream("test.xlsx")) {
            ExcelGenerator.generateExcel(fos, root, users, new HashMap<>());
        }
        System.out.println("엑셀 파일 생성 완료");

        // 4. 엑셀 파일 읽어서 DTO 변환 (headerRoot 넘겨줌)
        List<UserDTO> importedUsers;
        try (InputStream fis = new FileInputStream("test.xlsx")) {
            importedUsers = ExcelGenerator.parseExcelToDto(fis, UserDTO.class, 2, root);
        }

        // 5. 출력
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
