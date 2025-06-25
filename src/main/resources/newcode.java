package com.ul.rams.controller;

import java.io.*;
import java.util.*;
import java.util.regex.*;
import java.util.AbstractMap;
import java.util.Map.Entry;
import java.util.stream.Collectors;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.text.*;
import org.apache.pdfbox.pdmodel.interactive.form.*;

public class PdfDocFlatParser {

    public static void main(String[] args) throws IOException {
        String pdfPath = "src/main/resources/NZ_Adverse.pdf";
        try (PDDocument document = PDDocument.load(new File(pdfPath))) {
            System.out.println("[INFO] Starting extraction of document: " + pdfPath);
            LinkedHashMap<String, Object> extractedData = extractDataFromDocument(document);
            System.out.println("[INFO] Extraction complete. Outputting JSON:");
            printJson(extractedData);

            System.out.println("\n--- KEYS AND VALUES ---");
            extractedData.forEach((k, v) -> System.out.println(k + " : " + v));
        }
    }

    public static LinkedHashMap<String, Object> extractDataFromDocument(PDDocument document) throws IOException {
        LinkedHashMap<String, Object> result = new LinkedHashMap<>();
        Map<String, Boolean> mandatoryStatus = new HashMap<>();

        PDAcroForm acroForm = document.getDocumentCatalog().getAcroForm();
        if (acroForm != null) {
            for (PDField field : acroForm.getFields()) {
                Entry<String, Boolean> labelAndMandatory = extractLabelAndMandatory(field.getFullyQualifiedName());
                String key = buildNZKey(labelAndMandatory.getKey());
                String value = cleanValue(field.getValueAsString());
                boolean isMandatory = labelAndMandatory.getValue();
                mandatoryStatus.put(key, isMandatory);

                if (field instanceof PDCheckBox) {
                    boolean checked = ((PDCheckBox) field).isChecked();
                    if (isMandatory) {
                        Map<String, Object> valueObj = Map.of("value", checked, "mandatory", true);
                        result.put(key, valueObj);
                        System.out.println("[SYSOUT] AcroForm (mandatory checkbox): " + key + " = " + valueObj);
                    } else {
                        result.put(key, checked);
                        System.out.println("[SYSOUT] AcroForm (checkbox): " + key + " = " + checked);
                    }
                } else if (isValidValue(value, key)) {
                    if (isMandatory) {
                        Map<String, Object> valueObj = Map.of("value", value, "mandatory", true);
                        result.put(key, valueObj);
                        System.out.println("[SYSOUT] AcroForm (mandatory text): " + key + " = " + valueObj);
                    } else {
                        result.put(key, value);
                        System.out.println("[SYSOUT] AcroForm (text): " + key + " = " + value);
                    }
                }
            }
        }

        PDFTextStripper stripper = new PDFTextStripper();
        String text = stripper.getText(document);
        parseTextSections(text, result, mandatoryStatus);

        return flattenAndCleanResult(result, mandatoryStatus);
    }

    private static void parseTextSections(String text, Map<String, Object> result, Map<String, Boolean> mandatoryStatus) {
        String[] lines = text.split("\\r?\\n");
        String section = null, lastKey = null, currentSubgroup = null;
        boolean inNarrative = false;
        StringBuilder narrativeBuilder = new StringBuilder();

        Pattern sectionPattern = Pattern.compile("^\\d+\\.\\s*([^-:\\n]+)");
        Pattern combinedPattern = Pattern.compile("([A-Za-z0-9_ \\-/\\(\\)&\\[\\].,#*'’]+):\\s*([☒☑☐])" + "|"
                + "([☒☑☐])\\s*([^☒☑☐:\\n]+?)(?=\\s*[☒☑☐]|\\s+[A-Za-z0-9_\\-/\\(\\)&\\[\\].,#*'’ ]+?:|:|$)" + "|"
                + "([A-Za-z0-9_\\-/\\(\\)&\\[\\].,#*'’ ]+?):\\s*([^:]+?)(?=\\s+[A-Za-z0-9_\\-/\\(\\)&\\[\\].,#*'’ ]+?:|$)");

        Map<String, Integer> labelCountMap = new HashMap<>();
        boolean justSawEventProblemCheckboxes = false;
        Set<String> eventProblemCheckboxKeys = Set.of("Hcp", "Other Caregiver", "Patient", "N/A");

        for (int i = 0; i < lines.length; i++) {
            String line = lines[i].trim();
            System.out.println("Line: " + line);
            if (line.isEmpty()) continue;

            Matcher sectionMatcher = sectionPattern.matcher(line);
            if (sectionMatcher.find()) {
                flushNarrative(result, lastKey, narrativeBuilder, inNarrative, section, currentSubgroup, labelCountMap, mandatoryStatus);
                section = normalizeSectionName(sectionMatcher.group(1));
                currentSubgroup = null;
                inNarrative = false;
                lastKey = null;
                justSawEventProblemCheckboxes = false;
                System.out.println("[INFO] Section detected: " + section);
                continue;
            }

            if (isStrongHeading(line)) {
                flushNarrative(result, lastKey, narrativeBuilder, inNarrative, section, currentSubgroup, labelCountMap, mandatoryStatus);
                section = normalizeSectionName(line);
                currentSubgroup = null;
                inNarrative = false;
                lastKey = null;
                justSawEventProblemCheckboxes = false;
                System.out.println("[INFO] Custom heading detected: " + section);
                continue;
            }

            if (isTableBoundary(line)) {
                flushNarrative(result, lastKey, narrativeBuilder, inNarrative, section, currentSubgroup, labelCountMap, mandatoryStatus);
                inNarrative = false;
                lastKey = null;
                justSawEventProblemCheckboxes = false;
            }

            if (line.endsWith(":") && line.length() > 3 && !line.matches(".*:.*:.*")) {
                boolean foundGroup = false;
                if (i + 1 < lines.length) {
                    String nextLine = lines[i + 1].trim();
                    if (!nextLine.isEmpty() && (nextLine.matches(".*[☒☑☐].*") || nextLine.matches(".*?:\\s*.*"))) {
                        foundGroup = true;
                    }
                }
                if (foundGroup) {
                    currentSubgroup = normalizeSectionName(line.replace(":", "").trim());
                    System.out.println("[INFO] Subgroup detected: " + currentSubgroup);
                    continue;
                }
            }

            if (justSawEventProblemCheckboxes) {
                if (!looksLikeGroupingOrInstruction(line) && !isLikelyLabelOrHeader(line)
                        && !line.matches(".*[☒☑☐].*")) {
                    if ("Description Of The Clinical Event Problem"
                            .equalsIgnoreCase(section.replace("_", " ").trim())) {
                        String narrativeKey = buildNZKey(section, currentSubgroup, "Narrative");
                        result.put(narrativeKey, cleanValue(line));
                        System.out.println("[SYSOUT] Assigned clinical event narrative to key: " + narrativeKey + " => " + line);
                        justSawEventProblemCheckboxes = false;
                    }
                }
            }

            if (inNarrative && lastKey != null) {
                if (line.isEmpty()) continue;
                if (looksLikeGroupingOrInstruction(line) || isTableBoundary(line)) {
                    if (narrativeBuilder.length() > 0) {
                        flushNarrative(result, lastKey, narrativeBuilder, inNarrative, section, currentSubgroup, labelCountMap, mandatoryStatus);
                    }
                    inNarrative = false;
                    lastKey = null;
                    System.out.println("[DEBUG] Skipping heading/grouping/table-boundary in narrative: " + line);
                    continue;
                }
                if ((line.endsWith(":") && line.length() > 1)
                        || line.contains("Remedial Actions/Corrective Action/Preventive Action")) {
                    if (narrativeBuilder.length() > 0) {
                        flushNarrative(result, lastKey, narrativeBuilder, inNarrative, section, currentSubgroup, labelCountMap, mandatoryStatus);
                    }
                    Entry<String, Boolean> labelAndMandatory = extractLabelAndMandatory(line.replace(":", "").trim());
                    String newKeyBase = buildNZKey(section, currentSubgroup, labelAndMandatory.getKey());
                    String newKey = makeUniqueKey(newKeyBase, labelCountMap);
                    lastKey = newKey;
                    mandatoryStatus.put(lastKey, labelAndMandatory.getValue());
                    inNarrative = true;
                    System.out.println("[DEBUG] Narrative start for: " + lastKey + (labelAndMandatory.getValue() ? " [mandatory]" : ""));
                    continue;
                }
                Matcher m = Pattern.compile("^([A-Za-z0-9_\\-/\\(\\)&\\[\\].,#*'’ ]+?):\\s*(.*)$").matcher(line);
                if (m.find()) {
                    if (narrativeBuilder.length() > 0) {
                        flushNarrative(result, lastKey, narrativeBuilder, inNarrative, section, currentSubgroup, labelCountMap, mandatoryStatus);
                    }
                    String label = m.group(1).trim();
                    String value = cleanValue(m.group(2));
                    Entry<String, Boolean> labelAndMandatory = extractLabelAndMandatory(label);
                    String keyBase = buildNZKey(section, currentSubgroup, labelAndMandatory.getKey());
                    String key = makeUniqueKey(keyBase, labelCountMap);
                    mandatoryStatus.put(key, labelAndMandatory.getValue());
                    Object parsedValue = parsePossibleBooleanOrDate(value, labelAndMandatory.getKey());
                    if (keyBase.endsWith("_M_F")) {
                        keyBase = keyBase.replace("_M_F", "_Gender");
                        key = keyBase;
                        parsedValue = extractGender(value);
                        System.out.println("[SYSOUT] Gender key normalized: " + key + " = " + parsedValue);
                    }
                    if (labelAndMandatory.getValue()) {
                        Map<String, Object> valueObj = Map.of("value", parsedValue, "mandatory", true);
                        result.put(key, valueObj);
                        System.out.println("[SYSOUT] Label:Value (mandatory/narrative): " + key + " = " + valueObj);
                    } else {
                        result.put(key, parsedValue);
                        System.out.println("[SYSOUT] Label:Value (narrative): " + key + " = " + parsedValue);
                    }
                    lastKey = key;
                    inNarrative = false;
                    continue;
                }
                if (narrativeBuilder.length() > 0) narrativeBuilder.append(" ");
                narrativeBuilder.append(line);
                continue;
            }

            if (looksLikeGroupingOrInstruction(line)) {
                System.out.println("[DEBUG] Skipping grouping/instructional line: " + line);
                continue;
            }

            Matcher matcher = combinedPattern.matcher(line);
            boolean matchedAny = false;
            boolean allCheckBoxLine = true;
            Set<String> foundCheckboxLabels = new HashSet<>();
            while (matcher.find()) {
                matchedAny = true;
                if (matcher.group(1) != null && matcher.group(2) != null) {
                    String label = matcher.group(1);
                    String box = matcher.group(2);
                    boolean isChecked = box.equals("☒") || box.equals("☑");
                    String keyBase = buildNZKey(section, currentSubgroup, label.trim());
                    String key = makeUniqueKey(keyBase, labelCountMap);
                    result.put(key, isChecked);
                    mandatoryStatus.put(key, false);
                    lastKey = key;
                    System.out.println("[DEBUG] Checkbox (colon) detected: " + key + " = " + isChecked);
                    if (eventProblemCheckboxKeys.contains(label.trim())) foundCheckboxLabels.add(label.trim());
                    else allCheckBoxLine = false;
                } else if (matcher.group(3) != null && matcher.group(4) != null) {
                    String checkedBox = matcher.group(3);
                    String label = matcher.group(4);
                    boolean isChecked = checkedBox.equals("☒") || checkedBox.equals("☑");
                    String keyBase = buildNZKey(section, currentSubgroup, label.trim());
                    String key = makeUniqueKey(keyBase, labelCountMap);
                    result.put(key, isChecked);
                    mandatoryStatus.put(key, false);
                    lastKey = key;
                    System.out.println("[DEBUG] Checkbox (symbol) detected: " + key + " = " + isChecked);
                    if (eventProblemCheckboxKeys.contains(label.trim())) foundCheckboxLabels.add(label.trim());
                    else allCheckBoxLine = false;
                } else if (matcher.group(5) != null && matcher.group(6) != null) {
                    String label = matcher.group(5);
                    String rawValue = matcher.group(6);
                    if (looksLikeGroupingOrInstruction(label) || looksLikeGroupingOrInstruction(rawValue)) {
                        System.out.println("[DEBUG] Skipping label:value as one side looks like a heading/grouping: " + label + " : " + rawValue);
                        continue;
                    }
                    Entry<String, Boolean> labelAndMandatory = extractLabelAndMandatory(label.trim());
                    String keyBase = buildNZKey(section, currentSubgroup, labelAndMandatory.getKey());
                    String key = makeUniqueKey(keyBase, labelCountMap);
                    mandatoryStatus.put(key, labelAndMandatory.getValue());

                    String valueStr = cleanValue(rawValue.trim());
                    String[] valueParts = valueStr.split(" (?=[A-Z][a-z]+( [A-Z][a-z]+)*[/:])", 2);
                    Object value;
                    if (valueParts.length > 1 && looksLikeGroupingOrInstruction(valueParts[1])) {
                        value = removeTrailingGroupingText(valueParts[0].trim());
                        System.out.println("[DEBUG] Label:Value detected with trailing heading: " + key + " = " + value);
                        result.put(key, value);
                        String newKey = buildNZKey(section, currentSubgroup, valueParts[1].trim());
                        lastKey = makeUniqueKey(newKey, labelCountMap);
                        inNarrative = true;
                        System.out.println("[DEBUG] New heading detected after label:value: " + lastKey);
                        continue;
                    } else {
                        value = parsePossibleBooleanOrDate(valueStr, labelAndMandatory.getKey());
                        if (keyBase.endsWith("_M_F")) {
                            keyBase = keyBase.replace("_M_F", "_Gender");
                            key = keyBase;
                            value = extractGender(valueStr);
                            System.out.println("[SYSOUT] Gender key normalized: " + key + " = " + value);
                        }
                        if (value instanceof String) value = removeTrailingGroupingText((String) value);
                    }
                    System.out.println("[DEBUG] Label:Value detected: " + key + " = " + value);
                    if (isValidValue(value, key)) {
                        if (labelAndMandatory.getValue()) {
                            Map<String, Object> valueObj = Map.of("value", value, "mandatory", true);
                            System.out.println("[DEBUG] Label:Value extracted (mandatory): " + key + " => " + valueObj);
                            result.put(key, valueObj);
                        } else {
                            System.out.println("[DEBUG] Label:Value extracted: " + key + " => " + value);
                            result.put(key, value);
                        }
                        lastKey = key;
                    }
                }
            }

            if (matchedAny && allCheckBoxLine && section != null
                    && "Description Of The Clinical Event Problem".equalsIgnoreCase(section.replace("_", " ").trim())) {
                if (!foundCheckboxLabels.isEmpty()) {
                    justSawEventProblemCheckboxes = true;
                    System.out.println("[DEBUG] Detected block of event problem checkboxes");
                }
            } else if (matchedAny) {
                justSawEventProblemCheckboxes = false;
            }

            if (!matchedAny && line.endsWith(":") && line.length() > 3) {
                flushNarrative(result, lastKey, narrativeBuilder, inNarrative, section, currentSubgroup, labelCountMap, mandatoryStatus);
                Entry<String, Boolean> labelAndMandatory = extractLabelAndMandatory(line.replace(":", "").trim());
                String keyBase = buildNZKey(section, currentSubgroup, labelAndMandatory.getKey());
                if (keyBase.endsWith("_M_F")) keyBase = keyBase.replace("_M_F", "_Gender");
                String key = keyBase;
                lastKey = key;
                mandatoryStatus.put(key, labelAndMandatory.getValue());
                inNarrative = true;
                System.out.println("[DEBUG] Narrative start for: " + lastKey + (labelAndMandatory.getValue() ? " [mandatory]" : ""));
                continue;
            }

            if (!matchedAny && isTableBoundary(line) && i + 1 < lines.length) {
                String nextLine = lines[i + 1].trim();
                if (!nextLine.isEmpty() && !isTableBoundary(nextLine) && !looksLikeGroupingOrInstruction(nextLine)) {
                    Entry<String, Boolean> labelAndMandatory = extractLabelAndMandatory(line);
                    String keyBase;
                    if (line.toLowerCase().contains("wand number")) {
                        keyBase = buildNZKey(section, currentSubgroup, "Wand Number");
                        result.put(keyBase, cleanValue(nextLine));
                        mandatoryStatus.put(keyBase, false);
                        System.out.println("[SYSOUT] Writing Wand Number to key: " + keyBase + " = " + nextLine);
                    } else {
                        keyBase = buildNZKey(section, currentSubgroup, labelAndMandatory.getKey());
                        String key = makeUniqueKey(keyBase, labelCountMap);
                        Object value = parsePossibleBooleanOrDate(cleanValue(nextLine), labelAndMandatory.getKey());
                        result.put(key, value);
                        mandatoryStatus.put(key, labelAndMandatory.getValue());
                    }
                    i++;
                    continue;
                }
            }

            if (lastKey != null && !isLikelyLabelOrHeader(line) && !matchedAny && !inNarrative) {
                Object value = parsePossibleBooleanOrDate(cleanValue(line), lastKey);
                if (lastKey.endsWith("_M_F")) value = extractGender(value.toString());
                if (value instanceof String) value = removeTrailingGroupingText((String) value);
                if (isValidValue(value, lastKey)) {
                    if ("NZ_Description_Of_The_Clinical_Event_Problem_N_A".equals(lastKey)) {
                        String narrativeKey = "NZ_Description_Of_The_Clinical_Event_Problem_Narrative";
                        result.put(narrativeKey, value);
                        System.out.println("[SYSOUT] Assigned clinical event narrative to key: " + narrativeKey + " => " + value);
                        System.out.println("[SYSOUT] Skipped overwriting checkbox key: " + lastKey + " with value: " + value);
                    } else {
                        System.out.println("[DEBUG] Single-line narrative or value: " + lastKey + " => " + value);
                        result.put(lastKey, value);
                    }
                }
                lastKey = null;
                justSawEventProblemCheckboxes = false;
            }
        }
        flushNarrative(result, lastKey, narrativeBuilder, inNarrative, section, currentSubgroup, labelCountMap, mandatoryStatus);
    }

    private static String cleanValue(String value) {
        if (value == null) return null;
        String trimmed = value.trim();
        trimmed = trimmed.replaceAll("^\\([^\\)]*\\)\\s*", "");
        trimmed = removeTrailingGroupingText(trimmed);
        return trimmed.trim();
    }

    private static boolean isTableBoundary(String line) {
        String l = line.trim().toLowerCase();
        return l.startsWith("list of other devices involved") || l.startsWith("if other implants involved")
                || l.startsWith("mfr/sponsor aware of other similar events")
                || l.startsWith("country where these similar adverse events occurred")
                || l.startsWith("additional comments");
    }

    private static Entry<String, Boolean> extractLabelAndMandatory(String label) {
        boolean mandatory = false;
        if (label != null && label.contains("*")) {
            mandatory = true;
            label = label.replace("*", "").trim();
        }
        label = label.replaceAll("^[^A-Za-z0-9]+", "");
        String[] groupingWords = { "indicate", "select", "choose", "tick", "check", "please", "provide", "enter",
                "describe", "for details see", "see", "if the device", "is the device", "attach" };
        String lower = label.toLowerCase();
        for (String grp : groupingWords) {
            if (lower.startsWith(grp)) {
                label = label.substring(grp.length()).trim();
                lower = label.toLowerCase();
            }
        }
        String[] words = label.trim().split("\\s+|_");
        if (words.length > 3) {
            label = String.join(" ", Arrays.copyOfRange(words, words.length - 3, words.length));
        } else if (words.length > 0) {
            label = String.join(" ", words);
        }
        if (label.equalsIgnoreCase("M/F")) {
            label = "Gender";
        }
        return new AbstractMap.SimpleEntry<>(label, mandatory);
    }

    private static String normalizeSectionName(String raw) {
        if (raw == null) return "";
        String s = raw.replaceAll("\\b\\d+\\.?\\b", "").replaceAll("\\s{2,}", " ").trim();
        String[] splitters = { " if ", " indicate", ":", "-", " please", " select", " choose", " tick", " check" };
        for (String splitter : splitters) {
            int idx = s.toLowerCase().indexOf(splitter);
            if (idx > 0) s = s.substring(0, idx).trim();
        }
        return s;
    }

    private static String buildNZKey(String... parts) {
        return Arrays.stream(parts)
                .filter(Objects::nonNull)
                .map(PdfDocFlatParser::normalizeToPascalCase)
                .filter(s -> !s.isEmpty())
                .collect(Collectors.joining("_", "NZ_", ""));
    }

    private static String normalizeToPascalCase(String input) {
        if (input == null) return "";
        input = input.replaceAll("\\(.*?\\)", "");
        input = input.replaceAll("[^a-zA-Z0-9]", " ");
        String[] parts = input.trim().split("\\s+");
        StringBuilder sb = new StringBuilder();
        for (String part : parts) {
            if (part.isEmpty()) continue;
            sb.append(Character.toUpperCase(part.charAt(0)));
            if (part.length() > 1) sb.append(part.substring(1).toLowerCase());
            sb.append("_");
        }
        if (sb.length() > 0) sb.setLength(sb.length() - 1);
        return sb.toString();
    }

    private static Object parsePossibleBooleanOrDate(String s, String label) {
        if (s.equalsIgnoreCase("true") || s.equalsIgnoreCase("yes")) return true;
        if (s.equalsIgnoreCase("false") || s.equalsIgnoreCase("no")) return false;
        if (label != null && label.toLowerCase().contains("date")) {
            Matcher m = Pattern.compile("(\\d{1,2}/\\d{1,2}/\\d{4})").matcher(s);
            if (m.find()) return m.group(1);
        }
        return s;
    }

    private static String removeTrailingGroupingText(String value) {
        if (value == null) return null;
        String trimmed = value.trim();
        String[] groupings = { "specific device information", "device information", "patient information",
                "event information", "details", "narrative", "summary", "category", "type", "manufacturer", "model",
                "catalog", "other", "comments", "notes", "example", "section", "subsection", "grouping", "header",
                "heading", "description", "explanation" };
        for (String grouping : groupings) {
            String pattern = "(?i)\\b" + Pattern.quote(grouping) + "\\b\\.?$";
            trimmed = trimmed.replaceAll(pattern, "").trim();
        }
        return trimmed;
    }

    private static void flushNarrative(Map<String, Object> result, String lastKey, StringBuilder narrativeBuilder,
                                       boolean inNarrative, String section, String currentSubgroup, Map<String, Integer> labelCountMap, Map<String, Boolean> mandatoryStatus) {
        System.out.println("[DEBUG] flushNarrative called with: inNarrative=" + inNarrative + ", lastKey=" + lastKey
                + ", narrativeBuilder.length()=" + narrativeBuilder.length());
        if (inNarrative && lastKey != null && narrativeBuilder.length() > 0
                && isValidValue(narrativeBuilder.toString(), lastKey)) {
            String narrative = cleanValue(narrativeBuilder.toString());
            Boolean isMandatory = mandatoryStatus.getOrDefault(lastKey, false);
            if (lastKey.endsWith("_Gender")) {
                narrative = extractGender(narrative);
                System.out.println("[SYSOUT] (flushNarrative) Gender key normalized: " + lastKey + " = " + narrative);
            }
            Matcher trailingNumber = Pattern.compile("^(.*?)(\\s+)(\\d+)$").matcher(narrative);
            if (trailingNumber.matches()) {
                String mainText = trailingNumber.group(1).trim();
                String trailing = trailingNumber.group(3);
                Object oldVal = result.get(lastKey);
                if (oldVal instanceof Map && ((Map<?, ?>) oldVal).containsKey("mandatory")) {
                    Map<String, Object> valueObj = Map.of("value", mainText, "mandatory", true);
                    System.out.println("[DEBUG] Flushing narrative for (mandatory) " + lastKey + ": " + valueObj);
                    result.put(lastKey, valueObj);
                } else {
                    System.out.println("[DEBUG] Flushing narrative for " + lastKey + ": " + mainText);
                    result.put(lastKey, mainText);
                }
                String wandKey = buildNZKey(section, currentSubgroup, "Wand Number");
                System.out.println("[SYSOUT] (flushNarrative) Writing Wand Number to key: " + wandKey + " = " + trailing);
                result.put(wandKey, trailing);
            } else {
                Object oldVal = result.get(lastKey);
                if (oldVal instanceof Map && ((Map<?, ?>) oldVal).containsKey("mandatory")) {
                    Map<String, Object> valueObj = Map.of("value", narrative, "mandatory", true);
                    System.out.println("[DEBUG] Flushing narrative for (mandatory) " + lastKey + ": " + valueObj);
                    result.put(lastKey, valueObj);
                } else {
                    System.out.println("[DEBUG] Flushing narrative for " + lastKey + ": " + narrative);
                    result.put(lastKey, narrative);
                }
            }
        }
        narrativeBuilder.setLength(0);
    }

    private static boolean looksLikeGroupingOrInstruction(String line) {
        String lower = line.toLowerCase().trim();
        if (lower.matches("^(specific )?(device|patient|event|report|information|details|narrative|summary|category|type|manufacturer|model|serial|lot|catalog|brand|other|comments|notes|example|section|subsection|grouping|header|heading|description|explanation)[ .:-]*$")) {
            return true;
        }
        if (lower.startsWith("(") && lower.endsWith(")")) return true;
        if (lower.contains("indicate") || lower.contains("see definition") || lower.contains("category")
                || lower.contains("grouping") || lower.contains("instruction")
                || lower.contains("report category (see definitions")
                || lower.contains("both implant date and explant dates") || lower.startsWith("for details see")
                || lower.startsWith("please") || lower.startsWith("attach") || lower.startsWith("if the device")
                || lower.contains("resolution of event and outcomes") || lower.contains("patient focused")
                || lower.contains("specific device information")) return true;
        if (lower.matches("^\\(.*\\)$")) return true;
        return false;
    }

    private static boolean isStrongHeading(String line) {
        String l = line.trim();
        return l.matches("^(Remedial Actions/Corrective Action/Preventive Action|Other Reporting Information)$");
    }

    private static boolean isLikelyLabelOrHeader(String line) {
        return line.endsWith(":") || line.length() < 3 || line.equals(line.toUpperCase());
    }

    private static boolean isValidValue(Object value, String key) {
        if (value == null) return false;
        String val = value.toString().trim();
        if (val.isEmpty()) return false;
        String lower = val.toLowerCase();
        if (lower.equals("not applicable")) return false;
        if (lower.startsWith("(") && lower.endsWith(")")) return false;
        if (lower.startsWith("please submit") || lower.contains("submit an initial report")
                || lower.contains("submit a final report") || lower.startsWith("provide as much detail")
                || lower.contains("see guidance") || lower.startsWith("specify") || lower.contains("attach")
                || lower.startsWith("note:") || lower.startsWith("example:") || lower.startsWith("for example:")
                || lower.startsWith("email:") || lower.matches("^\\*?age:?$") || lower.matches("^\\*?wt.\\(kg\\):?$")
                || lower.matches("^\\*?m/f:?$") || lower.contains("guidance") || lower.equals("none")
                || lower.equals("click here to enter text") || lower.contains("send this form to")
                || lower.startsWith("if there have been other similar events reported")
                || lower.contains("if none, write") || lower.matches("^\\W*$")
                || lower.contains("the first report that the reporter")
                || lower.contains("submit this report when the investigation is complete")
                || lower.contains("number should include the number sold")
                || lower.contains("in some cases, the patient’s age") || lower.contains("incidence rate")
                || lower.contains("should preferably be provided in the form of an incidence rate")
                || lower.contains("this investigation should include details such as")
                || lower.contains("critical information that should be provided includes")
                || lower.contains("report types") || lower.contains("clinical event information")
                || lower.contains("manufacturer’s investigation") || lower.contains("harm definitions")
                || lower.contains("where required, to provide an update to a previous report")
                || lower.contains("report category") || lower.startsWith("●")
                || lower.contains("investigation is not yet complete and the final report not available."))
            return false;
        String lowerKey = key != null ? key.toLowerCase() : "";
        if (lowerKey.endsWith("note") || lowerKey.endsWith("example")) return false;
        if (val.length() < 2 && !val.equalsIgnoreCase("no")) return false;
        return true;
    }

    private static boolean isValidValue(Object value) {
        return isValidValue(value, "");
    }

    private static LinkedHashMap<String, Object> flattenAndCleanResult(Map<String, Object> map, Map<String, Boolean> mandatoryStatus) {
        LinkedHashMap<String, Object> cleaned = new LinkedHashMap<>();
        map.forEach((key, value) -> {
            Boolean isMandatory = mandatoryStatus.getOrDefault(key, false);
            if (isValidValue(value, key)) {
                if (value instanceof Map) {
                    Map<?, ?> vMap = (Map<?, ?>) value;
                    if (Boolean.TRUE.equals(vMap.get("mandatory"))) {
                        cleaned.put(key, value);
                    } else {
                        cleaned.put(key, vMap.get("value"));
                    }
                } else if (isMandatory) {
                    Map<String, Object> valueObj = Map.of("value", value, "mandatory", true);
                    cleaned.put(key, valueObj);
                } else {
                    cleaned.put(key, value);
                }
            }
        });
        return cleaned;
    }

    private static String makeUniqueKey(String base, Map<String, Integer> labelCountMap) {
        int count = labelCountMap.merge(base, 1, Integer::sum);
        return count == 1 ? base : base + "_" + count;
    }

    public static void printJson(Object obj) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        mapper.enable(SerializationFeature.INDENT_OUTPUT);
        String json = mapper.writeValueAsString(obj);
        System.out.println("\n--- FINAL JSON OUTPUT ---\n" + json);
    }

    private static String extractGender(String raw) {
        if (raw == null) return null;
        String lower = raw.trim().toLowerCase();
        if (lower.startsWith("male")) return "Male";
        if (lower.startsWith("female")) return "Female";
        if (lower.startsWith("other")) return "Other";
        String[] words = raw.trim().split("\\s+");
        if (words.length > 0) {
            String first = words[0].toLowerCase();
            if (first.equals("male") || first.equals("female") || first.equals("other")) {
                return Character.toUpperCase(first.charAt(0)) + first.substring(1).toLowerCase();
            }
        }
        return raw;
    }
}