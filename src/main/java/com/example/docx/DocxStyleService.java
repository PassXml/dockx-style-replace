package com.example.docx;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.Style;
import org.docx4j.wml.Styles;
import org.docx4j.XmlUtils;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Optional;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * 提供样式迁移、导出与清理的核心服务。
 */
public class DocxStyleService {

    public Path migrate(Path source, Path target, Set<String> styleNames, boolean copyNumbering, boolean includeDependencies) throws Exception {
        List<Path> tempFiles = new ArrayList<>();
        ConvertedFile src = ensureDocx(source, tempFiles, true);
        ConvertedFile dst = ensureDocx(target, tempFiles, false);
        try {
            transferStyles(src.effectivePath, dst.effectivePath, styleNames, copyNumbering, includeDependencies);
            return finalizeTarget(dst, target);
        } finally {
            cleanupTemps(tempFiles);
        }
    }

    public Path migrateAll(Path source, Path target, boolean copyNumbering) throws Exception {
        List<Path> tempFiles = new ArrayList<>();
        ConvertedFile src = ensureDocx(source, tempFiles, true);
        ConvertedFile dst = ensureDocx(target, tempFiles, false);
        try {
            transferAllStyles(src.effectivePath, dst.effectivePath, copyNumbering);
            return finalizeTarget(dst, target);
        } finally {
            cleanupTemps(tempFiles);
        }
    }

    public void exportStyles(Path source, Path output) throws Exception {
        List<Path> tempFiles = new ArrayList<>();
        ConvertedFile src = ensureDocx(source, tempFiles, true);
        try {
            dumpStyles(src.effectivePath, output);
        } finally {
            cleanupTemps(tempFiles);
        }
    }

    public CleanResult cleanStyles(Path document, Set<String> styleNames) throws Exception {
        if (styleNames.isEmpty()) {
            throw new IllegalArgumentException("至少需要一个样式名称或 ID 才能清理。");
        }
        List<Path> tempFiles = new ArrayList<>();
        ConvertedFile target = ensureDocx(document, tempFiles, false);
        try {
            int removed = removeStyles(target.effectivePath, styleNames);
            Path resultPath = finalizeTarget(target, document);
            return new CleanResult(resultPath, removed);
        } finally {
            cleanupTemps(tempFiles);
        }
    }

    private void transferStyles(Path srcDocx, Path dstDocx, Set<String> styleNames, boolean copyNumbering, boolean includeDependencies) throws Docx4JException {
        WordprocessingMLPackage srcPkg = WordprocessingMLPackage.load(srcDocx.toFile());
        WordprocessingMLPackage dstPkg = WordprocessingMLPackage.load(dstDocx.toFile());
        MainDocumentPart srcMain = srcPkg.getMainDocumentPart();
        MainDocumentPart dstMain = dstPkg.getMainDocumentPart();
        StyleDefinitionsPart srcStylePart = Optional.ofNullable(srcMain.getStyleDefinitionsPart())
                .orElseThrow(() -> new IllegalStateException("来源文件缺少样式定义"));
        StyleDefinitionsPart dstStylePart = dstMain.getStyleDefinitionsPart(true);

        Map<String, Style> stylesToCopy = collectStyles(srcStylePart, styleNames, includeDependencies);
        removeExistingStyles(dstStylePart, stylesToCopy.keySet());
        addStyles(dstStylePart, stylesToCopy);
        syncLatentAndDefaults(srcStylePart, dstStylePart);
        if (copyNumbering) {
            copyNumberingPart(srcMain, dstMain);
        }
        dstPkg.save(dstDocx.toFile());
    }

    private void transferAllStyles(Path srcDocx, Path dstDocx, boolean copyNumbering) throws Docx4JException {
        WordprocessingMLPackage srcPkg = WordprocessingMLPackage.load(srcDocx.toFile());
        WordprocessingMLPackage dstPkg = WordprocessingMLPackage.load(dstDocx.toFile());
        MainDocumentPart srcMain = srcPkg.getMainDocumentPart();
        MainDocumentPart dstMain = dstPkg.getMainDocumentPart();
        StyleDefinitionsPart srcStylePart = Optional.ofNullable(srcMain.getStyleDefinitionsPart())
                .orElseThrow(() -> new IllegalStateException("来源文件缺少样式定义"));
        StyleDefinitionsPart dstStylePart = dstMain.getStyleDefinitionsPart(true);

        List<Style> dstList = dstStylePart.getJaxbElement().getStyle();
        dstList.clear();
        for (Style style : srcStylePart.getJaxbElement().getStyle()) {
            dstList.add(XmlUtils.deepCopy(style));
        }
        syncLatentAndDefaults(srcStylePart, dstStylePart);
        if (copyNumbering) {
            copyNumberingPart(srcMain, dstMain);
        }
        dstPkg.save(dstDocx.toFile());
    }

    public List<StyleInfo> listStyles(Path document) throws Exception {
        List<Path> tempFiles = new ArrayList<>();
        ConvertedFile src = ensureDocx(document, tempFiles, true);
        try {
            return readStyles(src.effectivePath);
        } finally {
            cleanupTemps(tempFiles);
        }
    }

    private void dumpStyles(Path docxPath, Path output) throws Exception {
        List<StyleInfo> styles = readStyles(docxPath);
        Path normalized = output.toAbsolutePath().normalize();
        Path parent = normalized.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        try (BufferedWriter writer = Files.newBufferedWriter(normalized, StandardCharsets.UTF_8)) {
            writer.write("styleId,name,type");
            writer.newLine();
            for (StyleInfo style : styles) {
                writer.write(csv(style.styleId()));
                writer.write(',');
                writer.write(csv(style.name()));
                writer.write(',');
                writer.write(csv(style.type()));
                writer.newLine();
            }
        }
    }

    private List<StyleInfo> readStyles(Path docxPath) throws Exception {
        WordprocessingMLPackage pkg = WordprocessingMLPackage.load(docxPath.toFile());
        MainDocumentPart main = pkg.getMainDocumentPart();
        StyleDefinitionsPart stylePart = Optional.ofNullable(main.getStyleDefinitionsPart())
                .orElseThrow(() -> new IllegalStateException("来源文件缺少样式定义"));
        List<Style> styles = new ArrayList<>(stylePart.getJaxbElement().getStyle());
        styles.sort((a, b) -> {
            String idA = a.getStyleId() == null ? "" : a.getStyleId();
            String idB = b.getStyleId() == null ? "" : b.getStyleId();
            return idA.compareToIgnoreCase(idB);
        });
        List<StyleInfo> result = new ArrayList<>(styles.size());
        for (Style style : styles) {
            String id = Optional.ofNullable(style.getStyleId()).orElse("");
            String name = style.getName() != null ? style.getName().getVal() : "";
            String typeValue = style.getType() != null ? style.getType().toString() : "";
            result.add(new StyleInfo(id, name, typeValue));
        }
        return result;
    }

    private int removeStyles(Path docxPath, Set<String> styleNames) throws Exception {
        WordprocessingMLPackage pkg = WordprocessingMLPackage.load(docxPath.toFile());
        MainDocumentPart main = pkg.getMainDocumentPart();
        StyleDefinitionsPart stylePart = Optional.ofNullable(main.getStyleDefinitionsPart())
                .orElseThrow(() -> new IllegalStateException("文件缺少样式定义"));
        Set<String> normalized = normalizeStyleKeys(styleNames);
        List<Style> styles = stylePart.getJaxbElement().getStyle();
        int before = styles.size();
        styles.removeIf(style -> shouldRemoveStyle(style, normalized));
        int removed = before - styles.size();
        if (removed > 0) {
            pkg.save(docxPath.toFile());
        }
        return removed;
    }

    private Map<String, Style> collectStyles(StyleDefinitionsPart stylePart, Set<String> styleNames, boolean includeDependencies) {
        Map<String, Style> collector = new LinkedHashMap<>();
        for (String name : styleNames) {
            Style style = findStyle(stylePart, name)
                    .orElseThrow(() -> new IllegalArgumentException("找不到样式:" + name));
            addStyleWithDependencies(stylePart, collector, style, includeDependencies);
        }
        return collector;
    }

    private void addStyleWithDependencies(StyleDefinitionsPart part, Map<String, Style> collector, Style style, boolean includeDependencies) {
        if (collector.containsKey(style.getStyleId())) {
            return;
        }
        collector.put(style.getStyleId(), XmlUtils.deepCopy(style));
        if (includeDependencies && style.getBasedOn() != null && style.getBasedOn().getVal() != null) {
            findStyle(part, style.getBasedOn().getVal())
                    .ifPresent(parent -> addStyleWithDependencies(part, collector, parent, true));
        }
    }

    private Optional<Style> findStyle(StyleDefinitionsPart part, String key) {
        if (key == null) {
            return Optional.empty();
        }
        String normalized = key.trim();
        Style byId = part.getStyleById(normalized);
        if (byId != null) {
            return Optional.of(byId);
        }
        for (Style s : part.getJaxbElement().getStyle()) {
            if (s.getName() != null && normalized.equalsIgnoreCase(s.getName().getVal())) {
                return Optional.of(s);
            }
        }
        return Optional.empty();
    }

    private void removeExistingStyles(StyleDefinitionsPart dstStylePart, Set<String> styleIds) {
        List<Style> styles = dstStylePart.getJaxbElement().getStyle();
        styles.removeIf(s -> styleIds.contains(s.getStyleId()));
    }

    private void addStyles(StyleDefinitionsPart dstStylePart, Map<String, Style> stylesToCopy) {
        List<Style> targetList = dstStylePart.getJaxbElement().getStyle();
        targetList.addAll(stylesToCopy.values());
    }

    private void syncLatentAndDefaults(StyleDefinitionsPart src, StyleDefinitionsPart dst) {
        Styles srcStyles = src.getJaxbElement();
        Styles dstStyles = dst.getJaxbElement();
        if (srcStyles.getLatentStyles() != null) {
            dstStyles.setLatentStyles(XmlUtils.deepCopy(srcStyles.getLatentStyles()));
        }
        if (srcStyles.getDocDefaults() != null) {
            dstStyles.setDocDefaults(XmlUtils.deepCopy(srcStyles.getDocDefaults()));
        }
    }

    private void copyNumberingPart(MainDocumentPart srcMain, MainDocumentPart dstMain) throws Docx4JException {
        NumberingDefinitionsPart srcNum = srcMain.getNumberingDefinitionsPart();
        if (srcNum == null) {
            return;
        }
        NumberingDefinitionsPart dstNum = dstMain.getNumberingDefinitionsPart();
        if (dstNum == null) {
            dstNum = new NumberingDefinitionsPart();
            dstMain.addTargetPart(dstNum);
        }
        dstNum.setJaxbElement(XmlUtils.deepCopy(srcNum.getJaxbElement()));
    }

    private ConvertedFile ensureDocx(Path path, List<Path> tempFiles, boolean readOnly) throws IOException {
        String extension = FilenameUtils.getExtension(path.getFileName().toString()).toLowerCase(Locale.ROOT);
        if ("docx".equals(extension)) {
            return new ConvertedFile(path, path, !readOnly);
        }
        if (!"doc".equals(extension)) {
            throw new IllegalArgumentException("仅支援 .doc 或 .docx 文件: " + path);
        }
        if (readOnly) {
            Path temp = Files.createTempFile("poi-doc-convert-", ".docx");
            tempFiles.add(temp);
            convertDocToDocx(path, temp);
            return new ConvertedFile(path, temp, false);
        }
        Path docxOutput = deriveDocxSibling(path);
        convertDocToDocx(path, docxOutput);
        return new ConvertedFile(path, docxOutput, false);
    }

    private Path deriveDocxSibling(Path docPath) throws IOException {
        String baseName = FilenameUtils.getBaseName(docPath.getFileName().toString());
        Path parent = docPath.getParent();
        if (parent == null) {
            parent = Paths.get(".");
        }
        Path candidate = parent.resolve(baseName + ".docx").toAbsolutePath().normalize();
        int index = 1;
        while (Files.exists(candidate)) {
            candidate = parent.resolve(baseName + "-" + index + ".docx").toAbsolutePath().normalize();
            index++;
        }
        return candidate;
    }

    private void convertDocToDocx(Path docPath, Path docxPath) throws IOException {
        try (InputStream in = Files.newInputStream(docPath);
             HWPFDocument hwpf = new HWPFDocument(in);
             XWPFDocument xwpf = new XWPFDocument()) {
            buildParagraphs(hwpf, xwpf);
            try (OutputStream out = Files.newOutputStream(docxPath)) {
                xwpf.write(out);
            }
        }
    }

    private void buildParagraphs(HWPFDocument hwpf, XWPFDocument xwpf) {
        Range range = hwpf.getRange();
        List<Table> tables = new ArrayList<>();
        TableIterator iterator = new TableIterator(range);
        while (iterator.hasNext()) {
            tables.add(iterator.next());
        }
        int tableIndex = 0;
        for (int i = 0; i < range.numParagraphs(); i++) {
            Paragraph paragraph = range.getParagraph(i);
            Table currentTable = null;
            if (tableIndex < tables.size()) {
                Table candidate = tables.get(tableIndex);
                if (paragraph.getStartOffset() >= candidate.getStartOffset() && paragraph.getEndOffset() <= candidate.getEndOffset()) {
                    currentTable = candidate;
                }
            }
            if (currentTable != null) {
                appendTable(xwpf, currentTable);
                i += currentTable.numParagraphs() - 1;
                tableIndex++;
            } else {
                appendParagraph(xwpf.createParagraph(), paragraph);
            }
        }
    }

    private void appendParagraph(XWPFParagraph target, Paragraph source) {
        target.setAlignment(convertAlignment(source.getJustification()));
        for (int idx = 0; idx < source.numCharacterRuns(); idx++) {
            CharacterRun run = source.getCharacterRun(idx);
            String text = sanitize(run.text());
            if (text.isEmpty()) {
                continue;
            }
            XWPFRun xRun = target.createRun();
            xRun.setText(text);
            xRun.setBold(run.isBold());
            xRun.setItalic(run.isItalic());
            if (run.getUnderlineCode() != 0) {
                xRun.setUnderline(UnderlinePatterns.SINGLE);
            }
            if (run.getFontSize() > 0) {
                xRun.setFontSize(run.getFontSize() / 2);
            }
            if (run.getFontName() != null) {
                xRun.setFontFamily(run.getFontName());
            }
        }
    }

    private void appendTable(XWPFDocument document, Table table) {
        XWPFTable xTable = document.createTable();
        // docx 默认会先生成一列，若无资料需移除
        if (xTable.getNumberOfRows() > 0 && table.numRows() == 0) {
            xTable.removeRow(0);
        }
        for (int rowIndex = 0; rowIndex < table.numRows(); rowIndex++) {
            TableRow row = table.getRow(rowIndex);
            XWPFTableRow xRow = rowIndex == 0 && xTable.getNumberOfRows() == 1
                    ? xTable.getRow(0)
                    : xTable.createRow();
            ensureCellCount(xRow, row.numCells());
            for (int cellIndex = 0; cellIndex < row.numCells(); cellIndex++) {
                TableCell cell = row.getCell(cellIndex);
                XWPFTableCell xCell = xRow.getCell(cellIndex);
                xCell.removeParagraph(0);
                for (int p = 0; p < cell.numParagraphs(); p++) {
                    Paragraph para = cell.getParagraph(p);
                    XWPFParagraph xPara = xCell.addParagraph();
                    appendParagraph(xPara, para);
                }
            }
        }
    }

    private void ensureCellCount(XWPFTableRow row, int cells) {
        while (row.getTableCells().size() < cells) {
            row.addNewTableCell();
        }
        while (row.getTableCells().size() > cells) {
            row.getTableCells().remove(row.getTableCells().size() - 1);
        }
    }

    private ParagraphAlignment convertAlignment(int justification) {
        return switch (justification) {
            case 1 -> ParagraphAlignment.CENTER;
            case 2 -> ParagraphAlignment.RIGHT;
            case 3 -> ParagraphAlignment.BOTH;
            default -> ParagraphAlignment.LEFT;
        };
    }

    private void cleanupTemps(List<Path> tempFiles) {
        for (Path temp : tempFiles) {
            try {
                Files.deleteIfExists(temp);
            } catch (IOException ignore) {
            }
        }
    }

    private Path finalizeTarget(ConvertedFile converted, Path userTarget) throws IOException {
        if (converted.copyBackToOriginal) {
            if (!converted.effectivePath.equals(userTarget)) {
                Files.copy(converted.effectivePath, userTarget, StandardCopyOption.REPLACE_EXISTING);
            }
            return userTarget;
        }
        if (!converted.originalPath.equals(converted.effectivePath)) {
            return converted.effectivePath;
        }
        Files.copy(converted.effectivePath, userTarget, StandardCopyOption.REPLACE_EXISTING);
        return userTarget;
    }

    private Set<String> normalizeStyleKeys(Set<String> styleNames) {
        return styleNames.stream()
                .map(name -> name == null ? "" : name.trim().toLowerCase(Locale.ROOT))
                .filter(token -> !token.isEmpty())
                .collect(Collectors.toCollection(LinkedHashSet::new));
    }

    private boolean shouldRemoveStyle(Style style, Set<String> normalizedKeys) {
        if (normalizedKeys.isEmpty() || style == null) {
            return false;
        }
        if (style.getStyleId() != null && normalizedKeys.contains(style.getStyleId().trim().toLowerCase(Locale.ROOT))) {
            return true;
        }
        if (style.getName() != null && style.getName().getVal() != null) {
            String name = style.getName().getVal().trim().toLowerCase(Locale.ROOT);
            return normalizedKeys.contains(name);
        }
        return false;
    }

    private String sanitize(String text) {
        if (text == null) {
            return "";
        }
        String cleaned = text.replace("\r", "").replace("\7", "");
        return cleaned;
    }

    private String csv(String value) {
        String safe = value == null ? "" : value;
        if (safe.contains("\"") || safe.contains(",") || safe.contains("\n")) {
            safe = "\"" + safe.replace("\"", "\"\"") + "\"";
        }
        return safe;
    }

    public record StyleInfo(String styleId, String name, String type) {
    }

    public record CleanResult(Path file, int removed) {
    }

    private record ConvertedFile(Path originalPath, Path effectivePath, boolean copyBackToOriginal) {
    }
}
