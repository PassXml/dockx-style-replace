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
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.TableRowAlign;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPrGeneral;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import java.io.BufferedWriter;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
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
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;
import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

/**
 * 提供样式迁移、导出、清理与文档格式化的核心服务。
 *
 * <p>这个类同时用到了两套 Word 处理库：</p>
 * <p>1. docx4j：更适合直接复制/删除样式、编号等底层定义。</p>
 * <p>2. Apache POI：更适合遍历段落、表格、run，并按模板重排内容。</p>
 *
 * <p>因此可以把这个类理解成两层：</p>
 * <p>1. “样式定义层”负责迁移 styles.xml、numbering.xml 这类结构。</p>
 * <p>2. “内容格式层”负责把段落、表格、Markdown 渲染成目标模板的视觉效果。</p>
 */
public class DocxStyleService {

    private static final String WORDPROCESSING_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static final Pattern MARKDOWN_HEADING = Pattern.compile("^(#{1,6})\\s+(.*)$");
    private static final Pattern MARKDOWN_BULLET = Pattern.compile("^(\\s*)[-+*]\\s+(.*)$");
    private static final Pattern MARKDOWN_ORDERED = Pattern.compile("^(\\s*)(\\d+)\\.\\s+(.*)$");
    private static final Pattern MARKDOWN_TABLE_SEPARATOR = Pattern.compile(
        "^\\s*\\|?(\\s*:?-{3,}:?\\s*\\|)+\\s*:?-{3,}:?\\s*\\|?\\s*$");
    private static final Pattern MARKDOWN_INLINE = Pattern.compile(
        "(\\*\\*[^*]+\\*\\*|__[^_]+__|`[^`]+`|\\*[^*]+\\*|_[^_]+_)");

    public Path migrate(Path source, Path target, Set<String> styleNames, boolean copyNumbering,
        boolean includeDependencies) throws Exception {
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
            // 先抽取内置模板，再把模板中的样式定义整体复制到目标文档，
            // 最后基于“原始段落样式名 -> 模板样式 ID”的映射逐段重设格式。
            Path template = extractBuiltinTemplate(tempFiles);
            Map<String, String> originalStyleNames = readParagraphStyleNames(target.effectivePath);
            TemplateStyleSet templateStyles = loadTemplateStyleSet(template);
            transferAllStyles(template, target.effectivePath, true);
            applyDocumentFormat(target.effectivePath, template, originalStyleNames, templateStyles);
            return finalizeTarget(target, document);
        } finally {
            cleanupTemps(tempFiles);
        }
    }

    public Path convertMarkdown(Path markdownFile, Path templatePath, String documentTitle) throws Exception {
        String markdown = Files.readString(markdownFile, StandardCharsets.UTF_8);
        try (InputStream templateIn = Files.newInputStream(templatePath);
            XWPFDocument document = new XWPFDocument(templateIn)) {
            // Markdown 转换不是“纯文本导出”，而是借用模板文档作为样式容器，
            // 这样生成后的标题、正文、表格会直接落到模板已有样式上。
            TemplateStyleSet templateStyles = loadTemplateStyleSet(templatePath);
            renderMarkdownDocument(document, markdown, documentTitle, templateStyles);
            Path output = Files.createTempFile("markdown-convert-", ".docx");
            try (OutputStream out = Files.newOutputStream(output)) {
                document.write(out);
            }
            normalizeImageLayout(output);
            return output;
        }
    }

    private void transferStyles(Path srcDocx, Path dstDocx, Set<String> styleNames, boolean copyNumbering,
        boolean includeDependencies) throws Docx4JException {
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

    private void applyDocumentFormat(Path docxPath,
        Path templatePath,
        Map<String, String> originalStyleNames,
        TemplateStyleSet templateStyles) throws IOException {
        try (InputStream in = Files.newInputStream(docxPath);
            InputStream templateIn = Files.newInputStream(templatePath);
            XWPFDocument document = new XWPFDocument(in);
            XWPFDocument templateDocument = new XWPFDocument(templateIn)) {
            applyTemplateSectionProperties(document, templateDocument);
            formatBodyElements(document.getBodyElements(), originalStyleNames, templateStyles, false);
            try (OutputStream out = Files.newOutputStream(docxPath)) {
                document.write(out);
            }
        }
    }

    private void applyTemplateSectionProperties(XWPFDocument document, XWPFDocument templateDocument) {
        CTSectPr templateSectPr = templateDocument.getDocument().getBody().getSectPr();
        if (templateSectPr == null) {
            return;
        }
        document.getDocument().getBody().setSectPr((CTSectPr) templateSectPr.copy());
    }

    private void renderMarkdownDocument(XWPFDocument document,
        String markdown,
        String documentTitle,
        TemplateStyleSet templateStyles) {
        // 这里是一个很直接的“按行扫描”状态机：
        // 依次识别代码块、标题、表格、列表、引用和普通段落，
        // 每识别出一种块级结构，就把对应内容追加到 Word 文档中。
        clearDocumentBody(document);
        MarkdownNumbering numbering = ensureMarkdownNumbering(document);

        List<String> lines = normalizeMarkdown(markdown);
        boolean titleWritten = false;
        if (documentTitle != null && !documentTitle.isBlank()) {
            appendStyledParagraph(document, documentTitle.trim(), resolveTitleStyleId(templateStyles), false);
            titleWritten = true;
        }

        for (int index = 0; index < lines.size(); ) {
            String line = lines.get(index);
            if (line.isBlank()) {
                index++;
                continue;
            }

            if (isFenceStart(line)) {
                index = appendCodeBlock(document, lines, index, templateStyles);
                continue;
            }

            Matcher headingMatcher = MARKDOWN_HEADING.matcher(line);
            if (headingMatcher.matches()) {
                int level = headingMatcher.group(1).length();
                String text = headingMatcher.group(2).trim();
                boolean useTitle = level == 1 && !titleWritten;
                appendStyledParagraph(document, text, resolveHeadingStyleId(level, useTitle, templateStyles), false);
                titleWritten = titleWritten || useTitle;
                index++;
                continue;
            }

            if (isTableBlock(lines, index)) {
                index = appendTableBlock(document, lines, index, templateStyles);
                continue;
            }

            if (isListLine(line)) {
                index = appendListBlock(document, lines, index, templateStyles, numbering);
                continue;
            }

            if (line.trim().startsWith(">")) {
                index = appendQuoteBlock(document, lines, index, templateStyles);
                continue;
            }

            index = appendParagraphBlock(document, lines, index, templateStyles);
        }

        if (document.getBodyElements().isEmpty()) {
            appendStyledParagraph(document, "", templateStyles.normalStyleId(), true);
        }
    }

    private void clearDocumentBody(XWPFDocument document) {
        while (!document.getBodyElements().isEmpty()) {
            document.removeBodyElement(0);
        }
    }

    private List<String> normalizeMarkdown(String markdown) {
        String normalized = markdown == null ? "" : markdown.replace("\r\n", "\n").replace('\r', '\n');
        return Arrays.asList(normalized.split("\n", -1));
    }

    private boolean isFenceStart(String line) {
        return line != null && line.trim().startsWith("```");
    }

    private int appendCodeBlock(XWPFDocument document, List<String> lines, int startIndex,
        TemplateStyleSet templateStyles) {
        int index = startIndex + 1;
        while (index < lines.size() && !isFenceStart(lines.get(index))) {
            // 代码块不依赖模板中的专用“代码样式”，而是先使用正文段落，
            // 再通过 run 的等宽字体覆盖实现一个稳定的兜底效果。
            XWPFParagraph paragraph = createStyledParagraph(document, templateStyles.normalStyleId(), false);
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            for (XWPFRun run : paragraph.getRuns()) {
                run.setFontFamily("Courier New");
            }
            XWPFRun run = paragraph.createRun();
            run.setFontFamily("Courier New");
            run.setText(lines.get(index));
            index++;
        }
        return index < lines.size() ? index + 1 : index;
    }

    private boolean isTableBlock(List<String> lines, int index) {
        if (index + 1 >= lines.size()) {
            return false;
        }
        return looksLikeTableRow(lines.get(index)) && MARKDOWN_TABLE_SEPARATOR.matcher(lines.get(index + 1)).matches();
    }

    private boolean looksLikeTableRow(String line) {
        if (line == null || line.isBlank()) {
            return false;
        }
        int pipeCount = 0;
        for (char ch : line.toCharArray()) {
            if (ch == '|') {
                pipeCount++;
            }
        }
        return pipeCount >= 1;
    }

    private int appendTableBlock(XWPFDocument document, List<String> lines, int startIndex,
        TemplateStyleSet templateStyles) {
        List<List<String>> rows = new ArrayList<>();
        rows.add(parseTableRow(lines.get(startIndex)));
        int index = startIndex + 2;
        while (index < lines.size() && looksLikeTableRow(lines.get(index)) && !lines.get(index).isBlank()) {
            rows.add(parseTableRow(lines.get(index)));
            index++;
        }

        int columnCount = rows.stream().mapToInt(List::size).max().orElse(1);
        XWPFTable table = document.createTable(Math.max(rows.size(), 1), columnCount);
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            XWPFTableRow row = table.getRow(rowIndex);
            ensureCellCount(row, columnCount);
            List<String> cells = rows.get(rowIndex);
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                String cellText = columnIndex < cells.size() ? cells.get(columnIndex) : "";
                XWPFTableCell cell = row.getCell(columnIndex);
                if (cell.getParagraphs().isEmpty()) {
                    cell.addParagraph();
                }
                XWPFParagraph paragraph = cell.getParagraphs().get(0);
                clearParagraphContent(paragraph);
                appendInlineRuns(paragraph, cellText);
                for (int extra = cell.getParagraphs().size() - 1; extra >= 1; extra--) {
                    cell.removeParagraph(extra);
                }
            }
        }
        formatTable(table, templateStyles, resolveAvailablePageWidth(table));
        return index;
    }

    private List<String> parseTableRow(String line) {
        String normalized = line == null ? "" : line.trim();
        if (normalized.startsWith("|")) {
            normalized = normalized.substring(1);
        }
        if (normalized.endsWith("|")) {
            normalized = normalized.substring(0, normalized.length() - 1);
        }
        String[] parts = normalized.split("\\|", -1);
        List<String> cells = new ArrayList<>(parts.length);
        for (String part : parts) {
            cells.add(part.trim());
        }
        return cells;
    }

    private boolean isListLine(String line) {
        return MARKDOWN_BULLET.matcher(line).matches() || MARKDOWN_ORDERED.matcher(line).matches();
    }

    private int appendListBlock(XWPFDocument document,
        List<String> lines,
        int startIndex,
        TemplateStyleSet templateStyles,
        MarkdownNumbering numbering) {
        int index = startIndex;
        while (index < lines.size() && isListLine(lines.get(index))) {
            String line = lines.get(index);
            Matcher ordered = MARKDOWN_ORDERED.matcher(line);
            boolean orderedList = ordered.matches();
            String leadingWhitespace;
            String text;
            if (orderedList) {
                leadingWhitespace = ordered.group(1);
                text = ordered.group(3).trim();
            } else {
                Matcher bullet = MARKDOWN_BULLET.matcher(line);
                bullet.matches();
                leadingWhitespace = bullet.group(1);
                text = bullet.group(2).trim();
            }
            XWPFParagraph paragraph = createStyledParagraph(document, templateStyles.normalStyleId(), false);
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            // Markdown 列表最终仍要落到 Word 的编号体系里。
            // 这里通过 numId + ilvl 控制“是有序/无序列表”以及“缩进层级”。
            applyListNumbering(
                paragraph,
                orderedList ? numbering.orderedNumId() : numbering.bulletNumId(),
                resolveListLevel(leadingWhitespace)
            );
            appendInlineRuns(paragraph, text);
            index++;
        }
        return index;
    }

    private MarkdownNumbering ensureMarkdownNumbering(XWPFDocument document) {
        XWPFNumbering numbering = document.getNumbering();
        if (numbering == null) {
            numbering = document.createNumbering();
        }
        BigInteger bulletAbstractNumId = numbering.addAbstractNum(new XWPFAbstractNum(createListAbstractNum(false)));
        BigInteger bulletNumId = numbering.addNum(bulletAbstractNumId);
        BigInteger orderedAbstractNumId = numbering.addAbstractNum(new XWPFAbstractNum(createListAbstractNum(true)));
        BigInteger orderedNumId = numbering.addNum(orderedAbstractNumId);
        return new MarkdownNumbering(bulletNumId, orderedNumId);
    }

    private CTAbstractNum createListAbstractNum(boolean ordered) {
        CTAbstractNum abstractNum = CTAbstractNum.Factory.newInstance();
        abstractNum.setAbstractNumId(BigInteger.ZERO);
        for (int level = 0; level < 9; level++) {
            // Word 最多常见支持 9 级列表，这里预先一次性建好。
            CTLvl lvl = abstractNum.addNewLvl();
            lvl.setIlvl(BigInteger.valueOf(level));
            lvl.addNewStart().setVal(BigInteger.ONE);
            lvl.addNewNumFmt().setVal(ordered ? STNumberFormat.DECIMAL : STNumberFormat.BULLET);
            lvl.addNewLvlText().setVal(ordered ? "%" + (level + 1) + "." : "\u2022");
            lvl.addNewLvlJc().setVal(STJc.LEFT);
            CTPPrGeneral ppr = lvl.addNewPPr();
            CTInd ind = ppr.addNewInd();
            ind.setLeft(BigInteger.valueOf(720L * (level + 1)));
            ind.setHanging(BigInteger.valueOf(360));
        }
        return abstractNum;
    }

    private void applyListNumbering(XWPFParagraph paragraph, BigInteger numId, int level) {
        paragraph.setNumID(numId);
        paragraph.setIndentationFirstLine(0);
        paragraph.setIndentationLeft(720 * (level + 1));
        paragraph.setIndentationHanging(360);

        CTPPr ppr = paragraph.getCTP().isSetPPr() ? paragraph.getCTP().getPPr() : paragraph.getCTP().addNewPPr();
        CTNumPr numPr = ppr.isSetNumPr() ? ppr.getNumPr() : ppr.addNewNumPr();
        if (!numPr.isSetNumId()) {
            numPr.addNewNumId();
        }
        numPr.getNumId().setVal(numId);
        if (!numPr.isSetIlvl()) {
            numPr.addNewIlvl();
        }
        numPr.getIlvl().setVal(BigInteger.valueOf(level));
    }

    private int resolveListLevel(String leadingWhitespace) {
        if (leadingWhitespace == null || leadingWhitespace.isEmpty()) {
            return 0;
        }
        int spaces = 0;
        for (char ch : leadingWhitespace.toCharArray()) {
            if (ch == '\t') {
                spaces += 4;
            } else if (ch == ' ') {
                spaces++;
            }
        }
        return Math.min(spaces / 2, 8);
    }

    private int appendQuoteBlock(XWPFDocument document, List<String> lines, int startIndex,
        TemplateStyleSet templateStyles) {
        StringBuilder builder = new StringBuilder();
        int index = startIndex;
        while (index < lines.size()) {
            String line = lines.get(index);
            if (line.isBlank() || !line.trim().startsWith(">")) {
                break;
            }
            if (!builder.isEmpty()) {
                builder.append(' ');
            }
            builder.append(line.trim().substring(1).trim());
            index++;
        }
        XWPFParagraph paragraph = createStyledParagraph(document, templateStyles.normalStyleId(), false);
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun prefixRun = paragraph.createRun();
        prefixRun.setItalic(true);
        prefixRun.setText("引用：");
        appendInlineRuns(paragraph, builder.toString());
        return index;
    }

    private int appendParagraphBlock(XWPFDocument document, List<String> lines, int startIndex,
        TemplateStyleSet templateStyles) {
        StringBuilder builder = new StringBuilder();
        int index = startIndex;
        while (index < lines.size()) {
            String line = lines.get(index);
            if (line.isBlank() || isFenceStart(line) || isTableBlock(lines, index) || isListLine(line) || line.trim()
                .startsWith(">")) {
                break;
            }
            if (MARKDOWN_HEADING.matcher(line).matches()) {
                break;
            }
            if (!builder.isEmpty()) {
                builder.append(' ');
            }
            builder.append(line.trim());
            index++;
        }
        appendStyledParagraph(document, builder.toString(), templateStyles.normalStyleId(), true);
        return index;
    }

    private void appendStyledParagraph(XWPFDocument document, String text, String styleId, boolean firstLineIndent) {
        XWPFParagraph paragraph = createStyledParagraph(document, styleId, firstLineIndent);
        appendInlineRuns(paragraph, text);
    }

    private XWPFParagraph createStyledParagraph(XWPFDocument document, String styleId, boolean firstLineIndent) {
        XWPFParagraph paragraph = document.createParagraph();
        if (styleId != null && !styleId.isBlank()) {
            paragraph.setStyle(styleId);
        }
        if (!firstLineIndent) {
            paragraph.setIndentationFirstLine(0);
        }
        return paragraph;
    }

    private void appendInlineRuns(XWPFParagraph paragraph, String text) {
        if (text == null || text.isEmpty()) {
            if (paragraph.getRuns().isEmpty()) {
                paragraph.createRun().setText("");
            }
            return;
        }

        // 这里只处理一小部分常见行内语法，目标是“够用且稳定”，
        // 不是完整 Markdown 解析器。
        Matcher matcher = MARKDOWN_INLINE.matcher(text);
        int cursor = 0;
        while (matcher.find()) {
            if (matcher.start() > cursor) {
                paragraph.createRun().setText(text.substring(cursor, matcher.start()));
            }
            String token = matcher.group();
            if (token.startsWith("**") || token.startsWith("__")) {
                XWPFRun run = paragraph.createRun();
                run.setBold(true);
                run.setText(token.substring(2, token.length() - 2));
            } else if (token.startsWith("`")) {
                XWPFRun run = paragraph.createRun();
                run.setFontFamily("Courier New");
                run.setText(token.substring(1, token.length() - 1));
            } else {
                XWPFRun run = paragraph.createRun();
                run.setItalic(true);
                run.setText(token.substring(1, token.length() - 1));
            }
            cursor = matcher.end();
        }
        if (cursor < text.length()) {
            paragraph.createRun().setText(text.substring(cursor));
        }
    }

    private void clearParagraphContent(XWPFParagraph paragraph) {
        for (int index = paragraph.getRuns().size() - 1; index >= 0; index--) {
            paragraph.removeRun(index);
        }
    }

    private String resolveTitleStyleId(TemplateStyleSet templateStyles) {
        if (templateStyles.titleStyleId() != null && !templateStyles.titleStyleId().isBlank()) {
            return templateStyles.titleStyleId();
        }
        if (templateStyles.heading1StyleId() != null && !templateStyles.heading1StyleId().isBlank()) {
            return templateStyles.heading1StyleId();
        }
        return templateStyles.normalStyleId();
    }

    private String resolveHeadingStyleId(int level, boolean title, TemplateStyleSet templateStyles) {
        if (title) {
            return resolveTitleStyleId(templateStyles);
        }
        if (level <= 2 && templateStyles.heading1StyleId() != null && !templateStyles.heading1StyleId().isBlank()) {
            return templateStyles.heading1StyleId();
        }
        if (level == 3 && templateStyles.heading2StyleId() != null && !templateStyles.heading2StyleId().isBlank()) {
            return templateStyles.heading2StyleId();
        }
        if (level >= 4 && templateStyles.heading3StyleId() != null && !templateStyles.heading3StyleId().isBlank()) {
            return templateStyles.heading3StyleId();
        }
        return templateStyles.normalStyleId();
    }

    private void formatBodyElements(List<IBodyElement> bodyElements, Map<String, String> originalStyleNames,
        TemplateStyleSet templateStyles, boolean inTable) {
        for (IBodyElement bodyElement : bodyElements) {
            if (bodyElement instanceof XWPFParagraph paragraph) {
                if (!inTable) {
                    // 普通段落根据原文样式名重新映射到模板样式。
                    applyTemplateParagraphStyle(paragraph, originalStyleNames, templateStyles);
                }
                continue;
            }
            if (bodyElement instanceof XWPFTable table) {
                // 表格不只改表框，还会递归处理单元格内段落和嵌套表格。
                formatTable(table, templateStyles, resolveAvailablePageWidth(bodyElement));
            }
        }
    }

    private void formatTable(XWPFTable table, TemplateStyleSet templateStyles, int availablePageWidth) {
        // 表格格式化的主入口，可以把它理解成 4 个阶段：
        // 1. 清掉原表格样式，避免来源模板的表格主题继续干扰。
        // 2. 统一设置对齐、内边距、边框、列宽。
        // 3. 识别“视觉上的表头”到底跨了几行。
        // 4. 逐个单元格清理段落直设格式，再套用模板里的表头/表体样式。
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
        // 第一行如果存在纵向合并，表头在视觉上可能跨多行，
        // 所以这里不是简单返回 1，而是顺着 vertical merge 继续探测。
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
            // 这里不是按“物理列序号”直接取 cell，
            // 而是按逻辑列定位，再检查该列是否仍处于纵向合并的 continue 状态。
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
            // 某些单元格可能横向合并了多列，所以这里用 gridSpan 把“视觉列号”映射回真实 cell。
            if (logicalColumn >= currentColumn && logicalColumn < currentColumn + span) {
                return cell;
            }
            currentColumn += span;
        }
        return null;
    }

    private int getGridSpan(XWPFTableCell cell) {
        if (cell == null || cell.getCTTc() == null || cell.getCTTc().getTcPr() == null
            || cell.getCTTc().getTcPr().getGridSpan() == null) {
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
                // 单元格里的段落常常带有来源文档残留的边框、缩进、行距、run 直设格式。
                // 这里先“去直设”，再统一回到模板样式，避免看起来像是样式复制失败。
                clearParagraphBorders(paragraph);
                trimLeadingWhitespace(paragraph);
                paragraph.setSpacingBefore(0);
                paragraph.setSpacingAfter(0);
                paragraph.setSpacingBetween(1.0d);
                paragraph.setIndentationFirstLine(0);
                paragraph.setStyle(headerRow ? templateStyles.tableHeaderStyleId() : templateStyles.tableBodyStyleId());
                for (XWPFRun run : paragraph.getRuns()) {
                    clearRunDirectFormatting(run);
                }
                continue;
            }
            if (bodyElement instanceof XWPFTable nestedTable) {
                // 嵌套表格按同一套规则递归处理，否则外层表格正常、内层表格仍会保留旧格式。
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
        // 可用宽度 = 页面宽度 - 左右页边距。
        // 后续列宽分配不会直接用整页宽度，而是用这里算出来的内容区宽度。
        int pageWidth = asInt(sectionProperties.getPgSz().getW(), 11906);
        CTPageMar pageMargin = sectionProperties.getPgMar();
        int leftMargin = pageMargin == null ? 1800 : asInt(pageMargin.getLeft(), 1800);
        int rightMargin = pageMargin == null ? 1800 : asInt(pageMargin.getRight(), 1800);
        int available = pageWidth - leftMargin - rightMargin;
        return Math.max(available, 2400);
    }

    private void applyAutoFitWidth(XWPFTable table, int availablePageWidth) {
        int columnCount = resolvePreferredColumnCount(table);
        if (columnCount == 0) {
            return;
        }
        int[] contentUnits = new int[columnCount];
        for (XWPFTableRow row : table.getRows()) {
            int logicalColumn = 0;
            for (XWPFTableCell cell : row.getTableCells()) {
                int span = Math.max(getGridSpan(cell), 1);
                int cellUnits = estimateCellWidthUnits(cell);
                // 合并单元格的内容不能只压到第一列，否则会把某一列权重拉得过大。
                // 这里按跨度均摊到覆盖的逻辑列上，让首行合并表头对列宽估算更平滑。
                int perColumnUnits = Math.max(4, (int) Math.ceil((double) cellUnits / span));
                for (int offset = 0; offset < span && logicalColumn + offset < contentUnits.length; offset++) {
                    // 同一列取“最宽内容”的估算值，避免某一行很短就把整列压得过窄。
                    contentUnits[logicalColumn + offset] = Math.max(contentUnits[logicalColumn + offset], perColumnUnits);
                }
                logicalColumn += span;
            }
        }

        // 这里不是读取 Word 的真实渲染宽度，而是做一个启发式估算：
        // 文本越长、中文越多，列宽权重越大，再按页面可用宽度做比例分配。
        int[] widths = buildColumnWidths(contentUnits, availablePageWidth);
        // Word 表格宽度通常要同时写 tblW、tblGrid、tcW 三层信息，
        // 否则不同查看器里可能出现“表格总宽对了，但列宽没跟上”的情况。
        table.setWidth(Integer.toString(Math.min(sum(widths), availablePageWidth)));
        table.getCTTbl().getTblPr().addNewTblW().setType(STTblWidth.DXA);
        table.getCTTbl().getTblPr().getTblW().setW(BigInteger.valueOf(Math.min(sum(widths), availablePageWidth)));
        ensureTableGrid(table, widths);

        for (XWPFTableRow row : table.getRows()) {
            int logicalColumn = 0;
            for (XWPFTableCell cell : row.getTableCells()) {
                int span = Math.max(getGridSpan(cell), 1);
                setCellWidth(cell, sumWidths(widths, logicalColumn, span));
                logicalColumn += span;
            }
        }
    }

    private int resolvePreferredColumnCount(XWPFTable table) {
        List<XWPFTableRow> rows = table.getRows();
        if (rows.isEmpty()) {
            return 0;
        }
        XWPFTableRow firstRow = rows.get(0);
        // 如果首行存在横向合并，常见情况是它只是大类表头，真实列结构在第二行。
        // 这时优先参考第二行的逻辑列数，避免把整张表误判成“只有首行那几个大单元格”。
        if (hasHorizontalMerge(firstRow) && rows.size() > 1) {
            int secondRowColumns = countLogicalColumns(rows.get(1));
            if (secondRowColumns > 0) {
                return secondRowColumns;
            }
        }
        return rows.stream()
                .filter(Objects::nonNull)
                .mapToInt(this::countLogicalColumns)
                .max()
                .orElse(0);
    }

    private boolean hasHorizontalMerge(XWPFTableRow row) {
        if (row == null) {
            return false;
        }
        for (XWPFTableCell cell : row.getTableCells()) {
            if (getGridSpan(cell) > 1) {
                return true;
            }
        }
        return false;
    }

    private int countLogicalColumns(XWPFTableRow row) {
        if (row == null) {
            return 0;
        }
        int total = 0;
        for (XWPFTableCell cell : row.getTableCells()) {
            total += Math.max(getGridSpan(cell), 1);
        }
        return total;
    }

    private int[] buildColumnWidths(int[] contentUnits, int availablePageWidth) {
        int columnCount = contentUnits.length;
        int paddingPerColumn = 240;
        int minimumWidth = 720;
        int usableWidth = Math.max(availablePageWidth - paddingPerColumn * columnCount, minimumWidth * columnCount);
        int totalUnits = 0;
        for (int i = 0; i < contentUnits.length; i++) {
            // 每列先给一个最小内容权重，避免空列被算成 0 宽。
            contentUnits[i] = Math.max(contentUnits[i], 4);
            totalUnits += contentUnits[i];
        }

        int[] widths = new int[columnCount];
        int allocated = 0;
        for (int i = 0; i < columnCount; i++) {
            // 第一轮按内容权重做比例分配，同时给每列预留固定 padding。
            int proportional = Math.max(minimumWidth, usableWidth * contentUnits[i] / Math.max(totalUnits, 1));
            widths[i] = proportional + paddingPerColumn;
            allocated += widths[i];
        }

        if (allocated > availablePageWidth) {
            // 如果总宽超出页面，就按比例整体压缩，但仍保留最小列宽下限。
            double scale = (double) availablePageWidth / allocated;
            allocated = 0;
            for (int i = 0; i < columnCount; i++) {
                widths[i] = Math.max(minimumWidth, (int) Math.floor(widths[i] * scale));
                allocated += widths[i];
            }
        }

        if (allocated < availablePageWidth) {
            // 最后一列吃掉余量，避免因为整数除法导致表格右侧留出明显空白。
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
            // 这里用一个很粗但稳定的估算：
            // ASCII 按 1 单位，中文等宽字符按 2 单位，空白也算 1 单位。
            // 目标不是精确排版，而是让中英文混排时列宽分配更接近肉眼观感。
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
        // tblGrid 是“整张表的列定义”，可以理解成表格骨架。
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
        // tcW 是“单元格自己的宽度声明”，与 tblGrid 配合后兼容性更稳。
        cell.getCTTc().getTcPr().getTcW().setType(STTblWidth.DXA);
        cell.getCTTc().getTcPr().getTcW().setW(BigInteger.valueOf(width));
    }

    private int sumWidths(int[] widths, int startColumn, int span) {
        int total = 0;
        for (int index = 0; index < span && startColumn + index < widths.length; index++) {
            total += widths[startColumn + index];
        }
        return total;
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

    private void applyTemplateParagraphStyle(XWPFParagraph paragraph, Map<String, String> originalStyleNames,
        TemplateStyleSet templateStyles) {
        String styleId = resolveParagraphStyleId(paragraph, originalStyleNames, templateStyles);
        if (styleId == null || styleId.isBlank()) {
            styleId = templateStyles.normalStyleId();
        }
        if (isHeadingStyle(styleId, templateStyles)) {
            // 某些源文档标题段落会带直接编号。
            // 如果保留下来，套用模板标题样式后可能出现“标题前多出编号”的副作用。
            clearDirectParagraphNumbering(paragraph);
        }
        paragraph.setStyle(styleId);
    }

    private String resolveParagraphStyleId(XWPFParagraph paragraph, Map<String, String> originalStyleNames,
        TemplateStyleSet templateStyles) {
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

    private boolean isHeadingStyle(String styleId, TemplateStyleSet templateStyles) {
        if (styleId == null || styleId.isBlank()) {
            return false;
        }
        return styleId.equals(templateStyles.heading1StyleId())
            || styleId.equals(templateStyles.heading2StyleId())
            || styleId.equals(templateStyles.heading3StyleId());
    }

    private void clearDirectParagraphNumbering(XWPFParagraph paragraph) {
        if (paragraph == null || paragraph.getCTP() == null || paragraph.getCTP().getPPr() == null) {
            return;
        }
        CTPPr properties = paragraph.getCTP().getPPr();
        if (properties.isSetNumPr()) {
            properties.unsetNumPr();
        }
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
            // 这里约定“通过样式名称找模板样式”，而不是写死 styleId。
            // 原因是同一模板在不同环境里，styleId 可能变化，但样式名称更稳定。
            String normalStyleId = findRequiredStyleId(styleIdByName, "normal");
            String tableHeaderStyleId = findRequiredStyleId(styleIdByName, "表格-表头居中");
            String tableBodyStyleId = findRequiredStyleId(styleIdByName, "表格-内容居中", "表格文字", "表格内容");
            String titleStyleId = findOptionalStyleId(styleIdByName, "封面");
            String heading1StyleId = findOptionalStyleId(styleIdByName, "heading 1");
            String heading2StyleId = findOptionalStyleId(styleIdByName, "heading 2");
            String heading3StyleId = findOptionalStyleId(styleIdByName, "heading 3");
            return new TemplateStyleSet(
                styleIdByName,
                normalStyleId,
                tableHeaderStyleId,
                tableBodyStyleId,
                titleStyleId,
                heading1StyleId,
                heading2StyleId,
                heading3StyleId
            );
        } catch (Exception ex) {
            throw new IOException("读取模板样式失败", ex);
        }
    }

    private String findRequiredStyleId(Map<String, String> styleIdByName, String... styleNames) {
        for (String styleName : styleNames) {
            String styleId = styleIdByName.get(normalizeStyleName(styleName));
            if (styleId != null) {
                return styleId;
            }
        }
        throw new IllegalStateException("模板缺少样式: " + String.join(" / ", styleNames));
    }

    private String findOptionalStyleId(Map<String, String> styleIdByName, String styleName) {
        return styleIdByName.get(normalizeStyleName(styleName));
    }

    private void clearRunDirectFormatting(XWPFRun run) {
        if (run != null && run.getCTR() != null && run.getCTR().isSetRPr()) {
            run.getCTR().unsetRPr();
        }
    }

    private Map<String, Style> collectStyles(StyleDefinitionsPart stylePart, Set<String> styleNames,
        boolean includeDependencies) {
        Map<String, Style> collector = new LinkedHashMap<>();
        for (String name : styleNames) {
            Style style = findStyle(stylePart, name)
                .orElseThrow(() -> new IllegalArgumentException("找不到样式:" + name));
            addStyleWithDependencies(stylePart, collector, style, includeDependencies);
        }
        return collector;
    }

    private void addStyleWithDependencies(StyleDefinitionsPart part, Map<String, Style> collector, Style style,
        boolean includeDependencies) {
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
        // 除了显式样式列表，Word 还会依赖 latentStyles / docDefaults。
        // 不同步这两部分时，目标文档可能出现“样式已复制但显示仍不一致”的问题。
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

    private void normalizeImageLayout(Path docxPath) throws IOException {
        Path temp = Files.createTempFile("docx-image-layout-", ".docx");
        try (ZipInputStream zipIn = new ZipInputStream(Files.newInputStream(docxPath));
            ZipOutputStream zipOut = new ZipOutputStream(Files.newOutputStream(temp))) {
            // docx 本质上是 zip 包。
            // 这里直接改写内部 XML，把图片布局统一成更可控的 anchor 形式。
            // 学习点：
            // 1. Word 里的图片如果是 wp:inline，表示“嵌入文字行内”，通常不能做环绕。
            // 2. 想要“四周环绕型”，通常要改成 wp:anchor，并补上 wrapSquare / positionH / positionV 等节点。
            ZipEntry entry;
            while ((entry = zipIn.getNextEntry()) != null) {
                byte[] content = zipIn.readAllBytes();
                ZipEntry newEntry = new ZipEntry(entry.getName());
                zipOut.putNextEntry(newEntry);
                if (shouldNormalizeImageLayout(entry.getName())) {
                    zipOut.write(rewriteImageLayoutXml(content));
                } else {
                    zipOut.write(content);
                }
                zipOut.closeEntry();
                zipIn.closeEntry();
            }
        }
        Files.move(temp, docxPath, StandardCopyOption.REPLACE_EXISTING);
    }

    private boolean shouldNormalizeImageLayout(String entryName) {
        if (entryName == null) {
            return false;
        }
        return "word/document.xml".equals(entryName)
            || entryName.matches("word/header\\d+\\.xml")
            || entryName.matches("word/footer\\d+\\.xml")
            || "word/footnotes.xml".equals(entryName)
            || "word/endnotes.xml".equals(entryName);
    }

    private byte[] rewriteImageLayoutXml(byte[] xmlBytes) throws IOException {
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            factory.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);
            Document document = factory.newDocumentBuilder().parse(new ByteArrayInputStream(xmlBytes));
            // 分两步做归一化：
            // 1. inline 图片转成 anchor。
            // 2. 已经是 anchor 的图片也重新生成一次，覆盖掉来源文档中不一致的定位参数。
            boolean changed = normalizeInlinePictures(document);
            changed |= normalizeAnchoredPictures(document);
            if (!changed) {
                return xmlBytes;
            }
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            transformerFactory.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);
            Transformer transformer = transformerFactory.newTransformer();
            transformer.setOutputProperty(OutputKeys.ENCODING, StandardCharsets.UTF_8.name());
            transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "no");
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            transformer.transform(new DOMSource(document), new StreamResult(out));
            return out.toByteArray();
        } catch (Exception ex) {
            throw new IOException("调整图片布局失败", ex);
        }
    }

    private boolean normalizeInlinePictures(Document document) {
        List<Element> inlineElements = new ArrayList<>();
        NodeList nodes = document.getElementsByTagNameNS(WORDPROCESSING_DRAWING_NS, "inline");
        for (int index = 0; index < nodes.getLength(); index++) {
            Node node = nodes.item(index);
            if (node instanceof Element element) {
                inlineElements.add(element);
            }
        }
        for (Element inline : inlineElements) {
            // inline -> anchor：这是把“行内图片”变成“可环绕浮动图片”的关键一步。
            Element anchor = buildStandardAnchor(document, inline, false);
            inline.getParentNode().replaceChild(anchor, inline);
        }
        return !inlineElements.isEmpty();
    }

    private boolean normalizeAnchoredPictures(Document document) {
        List<Element> anchorElements = new ArrayList<>();
        NodeList nodes = document.getElementsByTagNameNS(WORDPROCESSING_DRAWING_NS, "anchor");
        for (int index = 0; index < nodes.getLength(); index++) {
            Node node = nodes.item(index);
            if (node instanceof Element element) {
                anchorElements.add(element);
            }
        }
        for (Element anchor : anchorElements) {
            // 已有 anchor 也要标准化，避免原文档残留其它环绕方式、偏移量、层级参数。
            Element normalized = buildStandardAnchor(document, anchor, true);
            anchor.getParentNode().replaceChild(normalized, anchor);
        }
        return !anchorElements.isEmpty();
    }

    private Element buildStandardAnchor(Document document, Element source, boolean preserveHorizontalPosition) {
        // 统一生成一套固定 anchor 结构，减少不同来源文档带来的布局漂移。
        // 可以把这里理解成在拼 Word 的“图片版式面板”：
        // - positionH / positionV：水平、垂直定位
        // - wrapSquare：四周环绕
        // - extent / graphic：图片尺寸与图形本体
        Element anchor = document.createElementNS(WORDPROCESSING_DRAWING_NS, "wp:anchor");
        anchor.setAttribute("distT", "0");
        anchor.setAttribute("distB", "0");
        anchor.setAttribute("distL", "0");
        anchor.setAttribute("distR", "0");
        anchor.setAttribute("simplePos", "0");
        anchor.setAttribute("relativeHeight", "0");
        anchor.setAttribute("behindDoc", "0");
        anchor.setAttribute("locked", "0");
        anchor.setAttribute("layoutInCell", "1");
        anchor.setAttribute("allowOverlap", "1");

        anchor.appendChild(createSimplePos(document));
        anchor.appendChild(createHorizontalPosition(document, source, preserveHorizontalPosition));
        anchor.appendChild(createVerticalPosition(document));
        appendClonedChildIfPresent(document, source, anchor, "extent");
        appendClonedChildIfPresent(document, source, anchor, "effectExtent");
        anchor.appendChild(createWrapSquare(document));
        appendClonedChildIfPresent(document, source, anchor, "docPr");
        appendClonedChildIfPresent(document, source, anchor, "cNvGraphicFramePr");
        appendClonedChildIfPresent(document, source, anchor, "graphic");
        return anchor;
    }

    private Element createSimplePos(Document document) {
        Element simplePos = document.createElementNS(WORDPROCESSING_DRAWING_NS, "wp:simplePos");
        simplePos.setAttribute("x", "0");
        simplePos.setAttribute("y", "0");
        return simplePos;
    }

    private Element createHorizontalPosition(Document document, Element source, boolean preserveHorizontalPosition) {
        if (preserveHorizontalPosition) {
            Element existing = findDirectChild(source, "positionH");
            if (existing != null) {
                // 水平位置允许沿用旧值，避免把用户原本左右摆放好的图片全部强制吸到同一列。
                return (Element) document.importNode(existing, true);
            }
        }
        Element positionH = document.createElementNS(WORDPROCESSING_DRAWING_NS, "wp:positionH");
        positionH.setAttribute("relativeFrom", "page");
        Element posOffset = document.createElementNS(WORDPROCESSING_DRAWING_NS, "wp:posOffset");
        posOffset.setTextContent("0");
        positionH.appendChild(posOffset);
        return positionH;
    }

    private Element createVerticalPosition(Document document) {
        Element positionV = document.createElementNS(WORDPROCESSING_DRAWING_NS, "wp:positionV");
        // relativeFrom="page" + posOffset=0
        // 对应 Word UI 里常见的“垂直：相对于页面，绝对位置 0”。
        positionV.setAttribute("relativeFrom", "page");
        Element posOffset = document.createElementNS(WORDPROCESSING_DRAWING_NS, "wp:posOffset");
        posOffset.setTextContent("0");
        positionV.appendChild(posOffset);
        return positionV;
    }

    private Element createWrapSquare(Document document) {
        Element wrapSquare = document.createElementNS(WORDPROCESSING_DRAWING_NS, "wp:wrapSquare");
        // wrapText="bothSides" 对应“四周环绕”，文字会在图片左右两侧绕排。
        wrapSquare.setAttribute("wrapText", "bothSides");
        return wrapSquare;
    }

    private void appendClonedChildIfPresent(Document document, Element source, Element target, String localName) {
        Element child = findDirectChild(source, localName);
        if (child != null) {
            target.appendChild(document.importNode(child, true));
        }
    }

    private Element findDirectChild(Element parent, String localName) {
        if (parent == null) {
            return null;
        }
        Node child = parent.getFirstChild();
        while (child != null) {
            if (child instanceof Element element
                && WORDPROCESSING_DRAWING_NS.equals(element.getNamespaceURI())
                && localName.equals(element.getLocalName())) {
                return element;
            }
            child = child.getNextSibling();
        }
        return null;
    }

    private ConvertedFile ensureDocx(Path path, List<Path> tempFiles, boolean readOnly) throws IOException {
        String extension = FilenameUtils.getExtension(path.getFileName().toString()).toLowerCase(Locale.ROOT);
        if ("docx".equals(extension)) {
            return new ConvertedFile(path, path, !readOnly);
        }
        if (!"doc".equals(extension)) {
            throw new IllegalArgumentException("仅支援 .doc 或 .docx 文件: " + path);
        }
        // 旧版 .doc 无法直接按 docx 逻辑处理，因此先转换成临时或旁路的 .docx。
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
        // HWPF 的段落流里“表格内容也是段落”，
        // 所以这里要一边扫段落，一边判断当前段落是否落在某个表格范围内。
        for (int i = 0; i < range.numParagraphs(); i++) {
            Paragraph paragraph = range.getParagraph(i);
            Table currentTable = null;
            if (tableIndex < tables.size()) {
                Table candidate = tables.get(tableIndex);
                if (paragraph.getStartOffset() >= candidate.getStartOffset()
                    && paragraph.getEndOffset() <= candidate.getEndOffset()) {
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
        normalizeImageLayout(converted.effectivePath);
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
        String tableBodyStyleId,
        String titleStyleId,
        String heading1StyleId,
        String heading2StyleId,
        String heading3StyleId
    ) {

    }

    private record MarkdownNumbering(BigInteger bulletNumId, BigInteger orderedNumId) {

    }
}
