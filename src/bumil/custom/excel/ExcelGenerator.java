package bumil.custom.excel;

import bumil.custom.excel.node.ExcelHeaderNode;
import bumil.custom.excel.node.ExcelStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;

import java.util.Map;
/**
 * 엑셀 파일을 생성 및 파싱하는 유틸리티 클래스입니다.
 * <p>다단계 헤더 지원, 셀 병합, 스타일 지정 기능을 포함합니다.</p>
 */
public class ExcelGenerator {
    /**
     * 엑셀 헤더를 렌더링하고 병합 처리를 수행합니다.
     *
     * @param sheet 워크시트 객체
     * @param node 현재 헤더 노드
     * @param rowIdx 시작 행 인덱스
     * @param colIdx 시작 열 인덱스
     * @param rowMap 행 캐시 맵
     * @param workbook 워크북 객체
     * @return 현재 노드가 차지하는 열의 수
     */
    private static int renderHeader(Sheet sheet, ExcelHeaderNode node, int rowIdx, int colIdx, Map<Integer, Row> rowMap, Workbook workbook) {
        Row row = rowMap.computeIfAbsent(rowIdx, sheet::createRow);
        Cell cell = row.createCell(colIdx);
        cell.setCellValue(node.getTitle());

        CellStyle cellStyle = createCellStyle(workbook, node.getStyle());
        cell.setCellStyle(cellStyle);

        if (node.getChildren() == null || node.getChildren().isEmpty()) {
            return 1; // leaf node = width 1
        }

        int width = 0;
        for (ExcelHeaderNode child : node.getChildren()) {
            width += renderHeader(sheet, child, rowIdx + 1, colIdx + width, rowMap, workbook);
        }

        if (width > 1) {
            sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, colIdx, colIdx + width - 1));
        } else {
            sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx + getDepth(node) - 1, colIdx, colIdx));
        }
        return width;
    }
    /**
     * 주어진 노드의 최대 깊이를 계산합니다.
     *
     * @param node 헤더 노드
     * @return 깊이 값
     */
    private static int getDepth(ExcelHeaderNode node) {
        if (node.getChildren() == null || node.getChildren().isEmpty()) return 1;
        return 1 + node.getChildren().stream().mapToInt(ExcelGenerator::getDepth).max().orElse(0);
    }
    /**
     * 리프 노드를 수집하여 리스트로 반환합니다.
     *
     * @param node 시작 노드
     * @return 리프 노드 리스트
     */
    private static List<ExcelHeaderNode> collectLeafNodes(ExcelHeaderNode node) {
        List<ExcelHeaderNode> leaves = new ArrayList<>();
        if (node.getChildren() == null || node.getChildren().isEmpty()) {
            leaves.add(node);
        } else {
            for (ExcelHeaderNode child : node.getChildren()) {
                leaves.addAll(collectLeafNodes(child));
            }
        }
        return leaves;
    }
    /**
     * DTO 객체의 필드 값을 반환합니다.
     *
     * @param dto 데이터 객체
     * @param fieldName 필드명
     * @return 필드 값
     * @throws Exception 리플렉션 오류 발생 시
     */
    private static Object getFieldValue(Object dto, String fieldName) throws Exception {
        Field field = dto.getClass().getDeclaredField(fieldName);
        field.setAccessible(true);
        return field.get(dto);
    }
    /**
     * DTO 객체에 필드 값을 설정합니다.
     *
     * @param dto 데이터 객체
     * @param fieldName 필드명
     * @param value 설정할 값
     * @throws Exception 리플렉션 오류 발생 시
     */
    private static void setFieldValue(Object dto, String fieldName, Object value) throws Exception {
        if (value == null) return;
        String setterName = "set" + Character.toUpperCase(fieldName.charAt(0)) + fieldName.substring(1);
        Method[] methods = dto.getClass().getMethods();
        for (Method method : methods) {
            if (method.getName().equals(setterName) && method.getParameterCount() == 1) {
                Class<?> paramType = method.getParameterTypes()[0];
                Object convertedValue = convertValue(value, paramType);
                method.invoke(dto, convertedValue);
                return;
            }
        }
    }
    /**
     * 타입에 맞게 값을 변환합니다.
     *
     * @param value 원본 값
     * @param targetType 대상 타입
     * @return 변환된 값
     */
    private static Object convertValue(Object value, Class<?> targetType) {
        if (targetType == Integer.class || targetType == int.class) {
            if (value instanceof Double) {
                return ((Double) value).intValue();
            }
        } else if (targetType == Long.class || targetType == long.class) {
            if (value instanceof Double) {
                return ((Double) value).longValue();
            }
        } else if (targetType == Double.class || targetType == double.class) {
            if (value instanceof Number) {
                return ((Number) value).doubleValue();
            }
        } else if (targetType == String.class) {
            return value.toString();
        } else if (targetType == Boolean.class || targetType == boolean.class) {
            if (value instanceof Boolean) {
                return value;
            }
        }
        return value;
    }
    /**
     * 주어진 스타일 정보를 기반으로 셀 스타일을 생성합니다.
     *
     * @param workbook 워크북 객체
     * @param excelStyle 스타일 정보
     * @return 생성된 셀 스타일
     */
    private static CellStyle createCellStyle(Workbook workbook, ExcelStyle excelStyle) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        if (excelStyle != null) {
            if (excelStyle.getFontName() != null) font.setFontName(excelStyle.getFontName());
            if (excelStyle.getFontSize() > 0) font.setFontHeightInPoints(excelStyle.getFontSize());
            font.setBold(excelStyle.isBold());
            style.setFont(font);
            if (excelStyle.getBackgroundColor() != null) {
                style.setFillForegroundColor(excelStyle.getBackgroundColor().getIndex());
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            if (excelStyle.getAlignment() != null) style.setAlignment(excelStyle.getAlignment());
            if (excelStyle.getVerticalAlignment() != null) style.setVerticalAlignment(excelStyle.getVerticalAlignment());
        }
        return style;
    }
    /**
     * DTO 리스트를 기반으로 엑셀 파일을 생성합니다.
     *
     * @param filename 생성할 파일명
     * @param headerRoot 헤더 트리 루트
     * @param dataList 데이터 객체 리스트
     * @param bodyStyleMap 필드명별 스타일 매핑
     * @throws Exception 파일 생성 실패 시
     */
    public static void generateExcel(String filename, ExcelHeaderNode headerRoot, List<?> dataList, Map<String, ExcelStyle> bodyStyleMap) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");
        Map<Integer, Row> rowMap = new HashMap<>();

        // 1. 헤더 렌더링
        renderHeader(sheet, headerRoot, 0, 0, rowMap, workbook);

        // 2. 리프 노드 수집
        List<ExcelHeaderNode> leafNodes = collectLeafNodes(headerRoot);

        // 3. 바디 렌더링
        int rowIdx = getDepth(headerRoot);
        for (Object dto : dataList) {
            Row row = sheet.createRow(rowIdx++);
            int colIdx = 0;
            for (ExcelHeaderNode leaf : leafNodes) {
                Cell cell = row.createCell(colIdx++);
                Object value = getFieldValue(dto, leaf.getFieldName());
                if (value != null) {
                    if (value instanceof Number) {
                        cell.setCellValue(((Number) value).doubleValue());
                    } else {
                        cell.setCellValue(value.toString());
                    }
                }

                // 스타일 적용
                ExcelStyle bodyStyle = bodyStyleMap.getOrDefault(leaf.getFieldName(), defaultBodyStyle(value));
                cell.setCellStyle(createCellStyle(workbook, bodyStyle));
            }
        }

        // 4. 파일 저장
        try (FileOutputStream fos = new FileOutputStream(filename)) {
            workbook.write(fos);
        }
        workbook.close();
    }

    private static ExcelStyle defaultBodyStyle(Object value) {
        ExcelStyle style = new ExcelStyle();
        style.setFontName("맑은 고딕");
        style.setFontSize((short) 10);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        if (value instanceof Number) {
            style.setAlignment(HorizontalAlignment.RIGHT);
        } else {
            style.setAlignment(HorizontalAlignment.CENTER);
        }

        return style;
    }
    /**
     * 엑셀 파일을 파싱하여 DTO 리스트로 변환합니다.
     *
     * @param filename 엑셀 파일명
     * @param dtoClass 변환할 DTO 클래스
     * @param headerEndRow 헤더 마지막 행 번호
     * @return 변환된 DTO 리스트
     * @throws Exception 파일 읽기 또는 매핑 실패 시
     */
    public static <T> List<T> parseExcelToDto(String filename, Class<T> dtoClass, int headerEndRow, ExcelHeaderNode headerRoot) throws Exception {
        List<T> result = new ArrayList<>();
        Map<String, String> titleToFieldNameMap = buildTitleToFieldNameMap(headerRoot);

        try (InputStream inputStream = new FileInputStream(filename)) {
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            Row headerRow = sheet.getRow(headerEndRow);
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) {
                headers.add(cell.getStringCellValue().trim());
            }

            for (int i = headerEndRow + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null || isEmptyRow(row)) continue;

                T dto = dtoClass.getDeclaredConstructor().newInstance();
                for (int j = 0; j < headers.size(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell == null) continue;

                    String fieldName = titleToFieldNameMap.get(headers.get(j)); // 한글 제목을 필드명으로 변환
                    if (fieldName == null) continue;

                    Object value = null;
                    switch (cell.getCellType()) {
                        case STRING: value = cell.getStringCellValue(); break;
                        case NUMERIC: value = cell.getNumericCellValue(); break;
                        case BOOLEAN: value = cell.getBooleanCellValue(); break;
                        default: break;
                    }
                    setFieldValue(dto, fieldName, value);
                }
                result.add(dto);
            }
            workbook.close();
        }
        return result;
    }
    /**
     * 주어진 행이 비어 있는지 여부를 검사합니다.
     *
     * @param row 검사할 행
     * @return 비어 있으면 true, 아니면 false
     */
    private static boolean isEmptyRow(Row row) {
        if (row == null) return true;
        for (Cell cell : row) {
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        return true;
    }

    // ExcelHeaderNode 트리로부터 title -> fieldName 매핑 생성
    private static Map<String, String> buildTitleToFieldNameMap(ExcelHeaderNode node) {
        Map<String, String> map = new HashMap<>();
        collectMapping(node, map);
        return map;
    }

    private static void collectMapping(ExcelHeaderNode node, Map<String, String> map) {
        if (node.getFieldName() != null) {
            map.put(node.getTitle(), node.getFieldName());
        }
        if (node.getChildren() != null) {
            for (ExcelHeaderNode child : node.getChildren()) {
                collectMapping(child, map);
            }
        }
    }
}
