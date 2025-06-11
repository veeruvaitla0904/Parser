package com.example.demo.parser;

import java.io.*;
import java.util.*;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.IRunElement;
import org.apache.poi.xwpf.usermodel.XWPFSDT;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

public class WordFormParser {

    public static void main(String[] args) {
        try (
                InputStream fis = WordFormParser.class.getClassLoader().getResourceAsStream("MDIR_Form.docx");
                XWPFDocument document = new XWPFDocument(fis)
        ) {
            LinkedHashMap<String, Object> extractedData = extractDataFromDocument(document);
            printJson(extractedData);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static LinkedHashMap<String, Object> extractDataFromDocument(XWPFDocument document) {
        List<IBodyElement> elements = document.getBodyElements();
        LinkedHashMap<String, Object> extractedData = new LinkedHashMap<>();
        List<String> currentHeadings = null;

        for (int i = 0; i < elements.size(); i++) {
            IBodyElement element = elements.get(i);

            if (element instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                List<String> headings = processParagraphForHeadings(paragraph);
                if (!headings.isEmpty()) {
                    currentHeadings = headings;
                }
            } else if (element instanceof XWPFTable) {
                XWPFTable table = (XWPFTable) element;
                if (currentHeadings != null && currentHeadings.size() > 1 && table.getNumberOfRows() == 1) {
                    // Special case: tab-separated headings + single-row table
                    processTableRowForMultipleHeadings(table.getRow(0), currentHeadings, extractedData);
                } else {
                    // Standard table extraction
                    String fallbackHeading = (currentHeadings != null && !currentHeadings.isEmpty())
                            ? currentHeadings.get(0)
                            : findFirstBoldCellText(table);
                    if (fallbackHeading == null) fallbackHeading = "Unnamed Section";
                    LinkedHashMap<String, Object> rowMap = extractTableData(table, fallbackHeading);
                    extractedData.put(fallbackHeading, rowMap);
                }
                currentHeadings = null;
            }
        }
        return flattenResult(extractedData);
    }

    private static List<String> processParagraphForHeadings(XWPFParagraph paragraph) {
        String text = paragraph.getText().trim();
        if (text.isEmpty()) return Collections.emptyList();
        if (text.equals("All fields marked with an * are mandatory fields (for the Final Report)")) {
            return Collections.emptyList();
        }
        for (XWPFRun run : paragraph.getRuns()) {
            if (run.isBold()) {
                List<String> headings = new ArrayList<>();
                String[] splits = text.split("\\t|\\r?\\n");
                for (String s : splits) {
                    String trimmed = s.trim();
                    if (!trimmed.isEmpty()) {
                        if (trimmed.endsWith(":")) {
                            trimmed = trimmed.substring(0, trimmed.length() - 1).trim();
                        }
                        headings.add(trimmed);
                    }
                }
                return headings;
            }
        }
        return Collections.emptyList();
    }

    private static String findFirstBoldCellText(XWPFTable table) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph para : cell.getParagraphs()) {
                    for (XWPFRun run : para.getRuns()) {
                        if (run.isBold()) {
                            String boldText = para.getText().trim();
                            if (!boldText.isEmpty()) {
                                return boldText.replaceAll(":$", "").trim();
                            }
                        }
                    }
                }
            }
        }
        return null;
    }

    private static List<String> extractAllSDTValues(XWPFTableCell cell) {
        List<String> values = new ArrayList<>();
        for (IBodyElement elem : cell.getBodyElements()) {
            if (elem instanceof XWPFSDT) {
                String val = ((XWPFSDT) elem).getContent().getText().trim();
                if (isValidValue(val)) values.add(val);
            }
        }
        for (XWPFParagraph para : cell.getParagraphs()) {
            for (IRunElement runElem : para.getIRuns()) {
                if (runElem instanceof XWPFSDT) {
                    String val = ((XWPFSDT) runElem).getContent().getText().trim();
                    if (isValidValue(val)) values.add(val);
                }
            }
            String paraText = para.getText().trim();
            if (isValidValue(paraText)) values.add(paraText);
        }
        Set<String> unique = new LinkedHashSet<>(values);
        return new ArrayList<>(unique);
    }

    private static void processTableRowForMultipleHeadings(XWPFTableRow row, List<String> headings,
                                                           LinkedHashMap<String, Object> outputMap) {
        List<String> sdtValues = new ArrayList<>();
        for (XWPFTableCell cell : row.getTableCells()) {
            sdtValues.addAll(extractAllSDTValues(cell));
        }
        for (int i = 0; i < headings.size(); i++) {
            String heading = headings.get(i);
            String value = (i < sdtValues.size()) ? sdtValues.get(i) : null;
            outputMap.put(heading, isValidValue(value) ? value : null);
        }
    }

    private static LinkedHashMap<String, Object> extractTableData(XWPFTable table, String currentHeading) {
        LinkedHashMap<String, Object> rowMap = new LinkedHashMap<>();
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                processTableCell(cell, rowMap, currentHeading);
            }
        }
        return rowMap;
    }

    private static void processTableCell(XWPFTableCell cell, LinkedHashMap<String, Object> rowMap,
                                         String currentHeading) {
        String rawText = cell.getText();
        if (rawText == null || rawText.trim().isEmpty()) return;
        String[] tabSplits = rawText.split("\t");
        for (String tabPart : tabSplits) {
            String[] lineSplits = tabPart.split("\\r?\\n");
            for (String line : lineSplits) {
                String text = line.trim();
                if (text.isEmpty()) continue;
                if (currentHeading != null && text.startsWith(currentHeading)) {
                    text = text.substring(currentHeading.length()).replaceFirst("^[:\\-]", "").trim();
                    if (text.isEmpty()) continue;
                }
                if (currentHeading != null && currentHeading.equalsIgnoreCase(text)) continue;
                boolean isMandatory = checkIfMandatory(cell, text);
                String key;
                String value = null;
                int colonIndex = text.indexOf(':');
                if (colonIndex != -1) {
                    key = text.substring(0, colonIndex).replaceAll("\\*$", "").trim();
                    value = text.substring(colonIndex + 1).trim();
                    if (value.isEmpty()) value = null;
                } else {
                    key = text.replaceAll("\\*$", "").trim();
                }
                String sdtValue = extractValueFromSDT(cell);
                if (isValidValue(sdtValue)) value = sdtValue;
                if (!isValidValue(value)) value = null;
                if (isMandatory) {
                    LinkedHashMap<String, Object> valObj = new LinkedHashMap<>();
                    valObj.put("value", value);
                    valObj.put("mandatory", true);
                    rowMap.put(key, valObj);
                } else {
                    rowMap.put(key, value);
                }
            }
        }
    }

    private static boolean checkIfMandatory(XWPFTableCell cell, String textToMatch) {
        for (XWPFParagraph para : cell.getParagraphs()) {
            String paraText = para.getText().trim();
            if (paraText.equals(textToMatch) && paraText.endsWith("*")) {
                return true;
            }
        }
        return false;
    }

    private static String extractValueFromSDT(XWPFTableCell cell) {
        for (IBodyElement elem : cell.getBodyElements()) {
            if (elem instanceof XWPFSDT) {
                String sdtValue = ((XWPFSDT) elem).getContent().getText().trim();
                if (isValidValue(sdtValue)) {
                    return sdtValue;
                }
            }
        }
        return null;
    }

    private static boolean isValidValue(String value) {
        if (value == null) return false;
        String val = value.trim();
        return !(val.isEmpty() || val.equalsIgnoreCase("Choose an item.")
                || val.equalsIgnoreCase("Click here to enter a date.")
                || val.equalsIgnoreCase("Choose")
                || val.equalsIgnoreCase("Blank")
                || val.equalsIgnoreCase("\"\"")
                || val.equals("*"));
    }

    private static LinkedHashMap<String, Object> flattenResult(LinkedHashMap<String, Object> sectionMap) {
        LinkedHashMap<String, Object> finalMap = new LinkedHashMap<>();
        for (Map.Entry<String, Object> entry : sectionMap.entrySet()) {
            String originalHeading = entry.getKey().trim();
            Object value = entry.getValue();
            boolean headingIsMandatory = originalHeading.endsWith("*") || originalHeading.endsWith(": *");
            // Remove trailing :, #, * and spaces
            String cleanedHeading = originalHeading.replaceAll("\\s*[:#\\*]+\\s*$", "").trim();
            if (value instanceof LinkedHashMap) {
                LinkedHashMap<?, ?> section = (LinkedHashMap<?, ?>) value;
                if (section.isEmpty()) {
                    if (headingIsMandatory) {
                        LinkedHashMap<String, Object> valObj = new LinkedHashMap<>();
                        valObj.put("value", null);
                        valObj.put("mandatory", true);
                        finalMap.put(cleanedHeading, valObj);
                    } else {
                        finalMap.put(cleanedHeading, null);
                    }
                } else {
                    LinkedHashMap<String, Object> cleanSection = new LinkedHashMap<>();
                    for (Map.Entry<?, ?> innerEntry : section.entrySet()) {
                        String k = innerEntry.getKey().toString().trim();
                        Object v = innerEntry.getValue();
                        // Remove trailing :, # and spaces from keys
                        k = k.replaceAll("\\s*[:#]+\\s*$", "").trim();
                        cleanSection.put(k, v);
                    }
                    finalMap.put(cleanedHeading, cleanSection);
                }
            } else {
                if (headingIsMandatory) {
                    LinkedHashMap<String, Object> valObj = new LinkedHashMap<>();
                    valObj.put("value", value);
                    valObj.put("mandatory", true);
                    finalMap.put(cleanedHeading, valObj);
                } else {
                    finalMap.put(cleanedHeading, value);
                }
            }
        }
        return finalMap;
    }

    private static void printJson(Object obj) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        mapper.enable(SerializationFeature.INDENT_OUTPUT);
        String json = mapper.writeValueAsString(obj);
        System.out.println(json);
    }
}