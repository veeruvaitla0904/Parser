import java.io.*;
import java.util.*;
import org.apache.poi.xwpf.usermodel.*;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

public class WordFormParserNZ {

    public static void main(String[] args) {
        try (
                InputStream fis = WordFormParserNZ.class.getClassLoader().getResourceAsStream("NZ_Adverse.docx");
                XWPFDocument document = new XWPFDocument(fis)
        ) {
            LinkedHashMap<String, Object> extractedData = extractDataFromDocument(document);
            printJson(extractedData);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static LinkedHashMap<String, Object> extractDataFromDocument(XWPFDocument document) {
        LinkedHashMap<String, Object> result = new LinkedHashMap<>();
        String currentMain = null;
        LinkedHashMap<String, Object> mainMap = null;
        String currentSub = null;
        LinkedHashMap<String, Object> subMap = null;
        boolean inMain = false, inSub = false, expectSubSection = false;

        List<XWPFTable> tables = document.getTables();
        if (tables == null || tables.isEmpty()) return result;
        XWPFTable table = tables.get(0);

        for (XWPFTableRow row : table.getRows()) {
            List<XWPFTableCell> cells = row.getTableCells();
            if (cells.isEmpty()) continue;
            XWPFTableCell cell = cells.get(0);
            String cellText = cell.getText().trim();
            if (cellText.isEmpty()) continue;

            boolean isMain = false, isBold = false;
            String color = null;
            for (XWPFParagraph para : cell.getParagraphs()) {
                for (XWPFRun run : para.getRuns()) {
                    if (run.isBold()) isBold = true;
                    if (run.getColor() != null) color = run.getColor();
                }
            }
            if (isBold && "FFFFFF".equalsIgnoreCase(color)) isMain = true;

            // Main section
            if (isMain) {
                if (currentSub != null && subMap != null && mainMap != null) {
                    mainMap.put(currentSub, subMap);
                }
                if (currentMain != null && mainMap != null) {
                    result.put(currentMain, mainMap);
                }
                currentMain = cellText.replaceAll(":$", "").trim();
                mainMap = new LinkedHashMap<>();
                currentSub = null;
                subMap = null;
                inMain = true;
                inSub = false;
                expectSubSection = true;
                continue;
            }

            // Sub-section (first bold after main)
            if (isBold && (color == null || "000000".equalsIgnoreCase(color)) && inMain && expectSubSection) {
                if (currentSub != null && subMap != null) {
                    mainMap.put(currentSub, subMap);
                }
                currentSub = cellText.replaceAll(":$", "").trim();
                subMap = new LinkedHashMap<>();
                inSub = true;
                expectSubSection = false;
                continue;
            }

            // Field under sub-section
            if (inSub && subMap != null && isBold && (color == null || "000000".equalsIgnoreCase(color))) {
                subMap.put(cellText.replaceAll(":$", "").trim(), null);
                continue;
            }

            // If another main section or sub-section is expected, reset
            if (isBold && (color == null || "000000".equalsIgnoreCase(color)) && inMain && !expectSubSection) {
                // This is a new sub-section
                if (currentSub != null && subMap != null) {
                    mainMap.put(currentSub, subMap);
                }
                currentSub = cellText.replaceAll(":$", "").trim();
                subMap = new LinkedHashMap<>();
                inSub = true;
                continue;
            }

            // Fallback: treat as field under main if not in sub
            if (inMain && mainMap != null && !inSub) {
                mainMap.put(cellText.replaceAll(":$", "").trim(), null);
            }
        }

        // Flush last sub-section and main section
        if (currentSub != null && subMap != null && mainMap != null) {
            mainMap.put(currentSub, subMap);
        }
        if (currentMain != null && mainMap != null) {
            result.put(currentMain, mainMap);
        }

        return result;
    }

    private static void printJson(Object obj) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        mapper.enable(SerializationFeature.INDENT_OUTPUT);
        String json = mapper.writeValueAsString(obj);
        System.out.println(json);
    }
}