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
import org.apache.poi.xwpf.usermodel.TableRowAlign;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.Style;
import org.docx4j.wml.Styles;
import org.docx4j.XmlUtils;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
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
import java.util.Objects;
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

    public Path formatDocument(Path document) throws Exception {
        List<Path> tempFiles = new ArrayList<>();
        ConvertedFile target = ensureDocx(document, tempFiles, false);
        try {
            Path template = extractBuiltinTemplate(tempFiles);
            Map<String, String> originalStyleNames = readParagraphStyleNames(target.effectivePath);
            TemplateStyleSet templateStyles = loadTemplateStyleSet(template);
            transferAllStyles(template, target.effectivePath, false);
            applyDocumentFormat(target.effectivePath, originalStyleNames, templateStyles);
            return finalizeTarget(target, document);
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

    private void applyDocumentFormat(Path docxPath, Map<String, String> originalStyleNames, TemplateStyleSet templateStyles) throws IOException {
        try (InputStream in = Files.newInputStream(docxPath);
             XWPFDocument document = new XWPFDocument(in)) {
            formatBodyElements(document.getBodyElements(), originalStyleNames, templateStyles, false);
            try (OutputStream out = Files.newOutputStream(docxPath)) {
                document.write(out);
            }
        }
    }

    private void formatBodyElements(List<IBodyElement> bodyElements, Map<String, String> originalStyleNames, TemplateStyleSet templateStyles, boolean inTable) {
        for (IBodyElement bodyElement : bodyElements) {
            if (bodyElement instanceof XWPFParagraph paragraph) {
                if (!inTable) {
                    applyTemplateParagraphStyle(paragraph, originalStyleNames, templateStyles);
                }
                continue;
            }
            if (bodyElement instanceof XWPFTable table) {
                formatTable(table, templateStyles, resolveAvailablePageWidth(bodyElement));
            }
        }
    }

    private void formatTable(XWPFTable table, TemplateStyleSet templateStyles, int availablePageWidth) {
        removeTableStyle(table);
        table.setTableAlignment(TableRowAlign.CENTER);
        table.setCellMargins(0, 108, 0, 108);
        applyAutoFitWidth(table, availablePageWidth);
        applyTableBorders(table);
        List<XWPFTableRow> rows = table.getRows();
        int headerRowCount = resolveHeaderRowCount(table);
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            XWPFTableRow row = rows.get(rowIndex);
            row.setHeight(360);
            boolean headerRow = rowIndex < headerRowCount;
            for (XWPFTableCell cell : row.getTableCells()) {
                cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                if (headerRow) {
                    cell.setColor("D9D9D9");
                }
                formatTableCell(cell, templateStyles, headerRow);
            }
        }
    }

    private int resolveHeaderRowCount(XWPFTable table) {
        List<XWPFTableRow> rows = table.getRows();
        if (rows.isEmpty()) {
            return 0;
        }
        int headerRowCount = 1;
        int logicalColumn = 0;
        for (XWPFTableCell cell : rows.get(0).getTableCells()) {
            int span = getGridSpan(cell);
            headerRowCount = Math.max(headerRowCount, resolveVerticalSpan(rows, logicalColumn));
            logicalColumn += span;
        }
        return headerRowCount;
    }

    private int resolveVerticalSpan(List<XWPFTableRow> rows, int logicalColumn) {
        int span = 1;
        for (int rowIndex = 1; rowIndex < rows.size(); rowIndex++) {
            XWPFTableCell cell = findCellByLogicalColumn(rows.get(rowIndex), logicalColumn);
            if (!isVerticalMergeContinuation(cell)) {
                break;
            }
            span++;
        }
        return span;
    }

    private XWPFTableCell findCellByLogicalColumn(XWPFTableRow row, int logicalColumn) {
        int currentColumn = 0;
        for (XWPFTableCell cell : row.getTableCells()) {
            int span = getGridSpan(cell);
            if (logicalColumn >= currentColumn && logicalColumn < currentColumn + span) {
                return cell;
            }
            currentColumn += span;
        }
        return null;
    }

    private int getGridSpan(XWPFTableCell cell) {
        if (cell == null || cell.getCTTc() == null || cell.getCTTc().getTcPr() == null || cell.getCTTc().getTcPr().getGridSpan() == null) {
            return 1;
        }
        return asInt(cell.getCTTc().getTcPr().getGridSpan().getVal(), 1);
    }

    private boolean isVerticalMergeContinuation(XWPFTableCell cell) {
        if (cell == null || cell.getCTTc() == null) {
            return false;
        }
        CTTcPr properties = cell.getCTTc().getTcPr();
        if (properties == null || properties.getVMerge() == null) {
            return false;
        }
        Object value = properties.getVMerge().getVal();
        if (value == null) {
            return true;
        }
        String normalized = value.toString().trim().toLowerCase(Locale.ROOT);
        return normalized.isEmpty() || "continue".equals(normalized);
    }

    private void formatTableCell(XWPFTableCell cell, TemplateStyleSet templateStyles, boolean headerRow) {
        for (IBodyElement bodyElement : cell.getBodyElements()) {
            if (bodyElement instanceof XWPFParagraph paragraph) {
                clearParagraphBorders(paragraph);
                trimLeadingWhitespace(paragraph);
                paragraph.setSpacingBefore(0);
                paragraph.setSpacingAfter(0);
                paragraph.setSpacingBetween(1.0d);
                paragraph.setIndentationFirstLine(0);
                paragraph.setAlignment(ParagraphAlignment.CENTER);
                paragraph.setStyle(headerRow ? templateStyles.tableHeaderStyleId() : templateStyles.tableBodyStyleId());
                for (XWPFRun run : paragraph.getRuns()) {
                    run.setFontFamily("宋体");
                    run.setFontSize(12);
                    run.setBold(headerRow);
                }
                continue;
            }
            if (bodyElement instanceof XWPFTable nestedTable) {
                formatTable(nestedTable, templateStyles, resolveAvailablePageWidth(nestedTable));
            }
        }
    }

    private int resolveAvailablePageWidth(IBodyElement bodyElement) {
        CTSectPr sectionProperties = null;
        if (bodyElement.getBody() instanceof XWPFDocument document) {
            sectionProperties = document.getDocument().getBody().getSectPr();
        }
        if (sectionProperties == null && bodyElement.getBody() instanceof XWPFTableCell cell) {
            sectionProperties = cell.getXWPFDocument().getDocument().getBody().getSectPr();
        }
        if (sectionProperties == null || sectionProperties.getPgSz() == null) {
            return 8266;
        }
        int pageWidth = asInt(sectionProperties.getPgSz().getW(), 11906);
        CTPageMar pageMargin = sectionProperties.getPgMar();
        int leftMargin = pageMargin == null ? 1800 : asInt(pageMargin.getLeft(), 1800);
        int rightMargin = pageMargin == null ? 1800 : asInt(pageMargin.getRight(), 1800);
        int available = pageWidth - leftMargin - rightMargin;
        return Math.max(available, 2400);
    }

    private void applyAutoFitWidth(XWPFTable table, int availablePageWidth) {
        int columnCount = table.getRows().stream()
                .filter(Objects::nonNull)
                .mapToInt(row -> row.getTableCells().size())
                .max()
                .orElse(0);
        if (columnCount == 0) {
            return;
        }
        int[] contentUnits = new int[columnCount];
        for (XWPFTableRow row : table.getRows()) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (int columnIndex = 0; columnIndex < cells.size(); columnIndex++) {
                int cellUnits = estimateCellWidthUnits(cells.get(columnIndex));
                contentUnits[columnIndex] = Math.max(contentUnits[columnIndex], cellUnits);
            }
        }

        int[] widths = buildColumnWidths(contentUnits, availablePageWidth);
        table.setWidth(Integer.toString(Math.min(sum(widths), availablePageWidth)));
        table.getCTTbl().getTblPr().addNewTblW().setType(STTblWidth.DXA);
        table.getCTTbl().getTblPr().getTblW().setW(BigInteger.valueOf(Math.min(sum(widths), availablePageWidth)));
        ensureTableGrid(table, widths);

        for (XWPFTableRow row : table.getRows()) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (int columnIndex = 0; columnIndex < cells.size() && columnIndex < widths.length; columnIndex++) {
                setCellWidth(cells.get(columnIndex), widths[columnIndex]);
            }
        }
    }

    private int[] buildColumnWidths(int[] contentUnits, int availablePageWidth) {
        int columnCount = contentUnits.length;
        int paddingPerColumn = 240;
        int minimumWidth = 720;
        int usableWidth = Math.max(availablePageWidth - paddingPerColumn * columnCount, minimumWidth * columnCount);
        int totalUnits = 0;
        for (int i = 0; i < contentUnits.length; i++) {
            contentUnits[i] = Math.max(contentUnits[i], 4);
            totalUnits += contentUnits[i];
        }

        int[] widths = new int[columnCount];
        int allocated = 0;
        for (int i = 0; i < columnCount; i++) {
            int proportional = Math.max(minimumWidth, usableWidth * contentUnits[i] / Math.max(totalUnits, 1));
            widths[i] = proportional + paddingPerColumn;
            allocated += widths[i];
        }

        if (allocated > availablePageWidth) {
            double scale = (double) availablePageWidth / allocated;
            allocated = 0;
            for (int i = 0; i < columnCount; i++) {
                widths[i] = Math.max(minimumWidth, (int) Math.floor(widths[i] * scale));
                allocated += widths[i];
            }
        }

        if (allocated < availablePageWidth) {
            widths[columnCount - 1] += availablePageWidth - allocated;
        }
        return widths;
    }

    private int estimateCellWidthUnits(XWPFTableCell cell) {
        int maxUnits = 4;
        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            int paragraphUnits = estimateTextWidthUnits(paragraph.getText());
            maxUnits = Math.max(maxUnits, paragraphUnits);
        }
        return maxUnits;
    }

    private int estimateTextWidthUnits(String text) {
        if (text == null || text.isBlank()) {
            return 4;
        }
        int units = 0;
        for (char ch : text.toCharArray()) {
            if (Character.isWhitespace(ch)) {
                units += 1;
            } else if (ch <= 127) {
                units += 1;
            } else {
                units += 2;
            }
        }
        return Math.max(units, 4);
    }

    private void ensureTableGrid(XWPFTable table, int[] widths) {
        if (table.getCTTbl().getTblGrid() == null) {
            table.getCTTbl().addNewTblGrid();
        } else {
            table.getCTTbl().getTblGrid().setGridColArray(null);
        }
        for (int width : widths) {
            table.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(width));
        }
    }

    private void setCellWidth(XWPFTableCell cell, int width) {
        if (cell.getCTTc().getTcPr() == null) {
            cell.getCTTc().addNewTcPr();
        }
        if (cell.getCTTc().getTcPr().getTcW() == null) {
            cell.getCTTc().getTcPr().addNewTcW();
        }
        cell.getCTTc().getTcPr().getTcW().setType(STTblWidth.DXA);
        cell.getCTTc().getTcPr().getTcW().setW(BigInteger.valueOf(width));
    }

    private int sum(int[] values) {
        int total = 0;
        for (int value : values) {
            total += value;
        }
        return total;
    }

    private int asInt(Object value, int defaultValue) {
        if (value instanceof BigInteger bigInteger) {
            return bigInteger.intValue();
        }
        if (value instanceof Number number) {
            return number.intValue();
        }
        if (value != null) {
            try {
                return Integer.parseInt(value.toString());
            } catch (NumberFormatException ignore) {
            }
        }
        return defaultValue;
    }

    private void applyTemplateParagraphStyle(XWPFParagraph paragraph, Map<String, String> originalStyleNames, TemplateStyleSet templateStyles) {
        String styleId = resolveParagraphStyleId(paragraph, originalStyleNames, templateStyles);
        if (styleId == null || styleId.isBlank()) {
            styleId = templateStyles.normalStyleId();
        }
        paragraph.setStyle(styleId);
    }

    private String resolveParagraphStyleId(XWPFParagraph paragraph, Map<String, String> originalStyleNames, TemplateStyleSet templateStyles) {
        String originalStyleId = paragraph.getStyle();
        String styleName = normalizeStyleName(originalStyleNames.get(originalStyleId));
        if (styleName != null) {
            String directMatch = templateStyles.styleIdByName().get(styleName);
            if (directMatch != null) {
                return directMatch;
            }
        }
        Integer headingLevel = detectHeadingLevel(originalStyleId, styleName, paragraph);
        if (headingLevel != null) {
            String headingStyle = templateStyles.styleIdByName().get("heading " + headingLevel);
            if (headingStyle != null) {
                return headingStyle;
            }
        }
        return templateStyles.normalStyleId();
    }

    private Integer detectHeadingLevel(String originalStyleId, String styleName, XWPFParagraph paragraph) {
        Integer fromName = extractHeadingLevel(styleName);
        if (fromName != null) {
            return fromName;
        }
        Integer fromId = extractHeadingLevel(normalizeStyleName(originalStyleId));
        if (fromId != null) {
            return fromId;
        }
        if (paragraph.getCTP().getPPr() != null && paragraph.getCTP().getPPr().getOutlineLvl() != null) {
            return paragraph.getCTP().getPPr().getOutlineLvl().getVal().intValue() + 1;
        }
        return null;
    }

    private Integer extractHeadingLevel(String value) {
        if (value == null) {
            return null;
        }
        if (value.startsWith("heading ")) {
            String suffix = value.substring("heading ".length()).trim();
            if (suffix.length() == 1 && Character.isDigit(suffix.charAt(0))) {
                return suffix.charAt(0) - '0';
            }
        }
        return null;
    }

    private String normalizeStyleName(String value) {
        if (value == null || value.isBlank()) {
            return null;
        }
        return value.trim().toLowerCase(Locale.ROOT);
    }

    private void removeTableStyle(XWPFTable table) {
        CTTblPr tableProperties = table.getCTTbl().getTblPr();
        if (tableProperties != null && tableProperties.isSetTblStyle()) {
            tableProperties.unsetTblStyle();
        }
    }

    private void applyTableBorders(XWPFTable table) {
        table.setTopBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        table.setBottomBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        table.setLeftBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        table.setRightBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        table.setInsideHBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        table.setInsideVBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
    }

    private void clearParagraphBorders(XWPFParagraph paragraph) {
        if (paragraph.getCTP().getPPr() != null && paragraph.getCTP().getPPr().isSetPBdr()) {
            paragraph.getCTP().getPPr().unsetPBdr();
        }
    }

    private void trimLeadingWhitespace(XWPFParagraph paragraph) {
        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.text();
            if (text == null || text.isEmpty()) {
                continue;
            }
            String cleaned = text.replaceFirst("^[\\t\\u00A0 ]+", "");
            if (!cleaned.equals(text)) {
                run.setText(cleaned, 0);
            }
            if (!cleaned.isBlank()) {
                return;
            }
        }
    }

    private Path extractBuiltinTemplate(List<Path> tempFiles) throws IOException {
        try (InputStream stream = DocxStyleService.class.getResourceAsStream("/template/template.docx")) {
            if (stream == null) {
                throw new IllegalStateException("找不到内置模板: /template/template.docx");
            }
            Path temp = Files.createTempFile("docx-format-template-", ".docx");
            Files.copy(stream, temp, StandardCopyOption.REPLACE_EXISTING);
            tempFiles.add(temp);
            return temp;
        }
    }

    private Map<String, String> readParagraphStyleNames(Path docxPath) throws IOException {
        try {
            Map<String, String> styleNames = new LinkedHashMap<>();
            for (StyleInfo styleInfo : readStyles(docxPath)) {
                if (!styleInfo.styleId().isBlank() && !styleInfo.name().isBlank()) {
                    styleNames.put(styleInfo.styleId(), styleInfo.name());
                }
            }
            return styleNames;
        } catch (Exception ex) {
            throw new IOException("读取文档样式失败", ex);
        }
    }

    private TemplateStyleSet loadTemplateStyleSet(Path templatePath) throws IOException {
        try {
            Map<String, String> styleIdByName = new LinkedHashMap<>();
            for (StyleInfo styleInfo : readStyles(templatePath)) {
                if (!styleInfo.styleId().isBlank() && !styleInfo.name().isBlank()) {
                    styleIdByName.put(normalizeStyleName(styleInfo.name()), styleInfo.styleId());
                }
            }
            String normalStyleId = findRequiredStyleId(styleIdByName, "normal");
            String tableHeaderStyleId = findRequiredStyleId(styleIdByName, "表格-表头居中");
            String tableBodyStyleId = findRequiredStyleId(styleIdByName, "表格内容");
            return new TemplateStyleSet(styleIdByName, normalStyleId, tableHeaderStyleId, tableBodyStyleId);
        } catch (Exception ex) {
            throw new IOException("读取模板样式失败", ex);
        }
    }

    private String findRequiredStyleId(Map<String, String> styleIdByName, String styleName) {
        String styleId = styleIdByName.get(normalizeStyleName(styleName));
        if (styleId == null) {
            throw new IllegalStateException("模板缺少样式: " + styleName);
        }
        return styleId;
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

    private record TemplateStyleSet(
            Map<String, String> styleIdByName,
            String normalStyleId,
            String tableHeaderStyleId,
            String tableBodyStyleId
    ) {
    }
}
