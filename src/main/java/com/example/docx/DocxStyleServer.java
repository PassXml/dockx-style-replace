package com.example.docx;

import io.javalin.Javalin;
import io.javalin.http.BadRequestResponse;
import io.javalin.http.Context;
import io.javalin.http.UploadedFile;
import org.apache.commons.io.FilenameUtils;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.LinkedHashMap;
import java.util.HashSet;
import java.util.HashMap;
import java.util.function.Function;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * 使用 Javalin 提供同步的 Web API。
 */
public class DocxStyleServer {

    private final DocxStyleService service = new DocxStyleService();
    private final Path workDir;

    public DocxStyleServer(Path workDir) throws IOException {
        this.workDir = Files.createDirectories(workDir);
    }

    public static void main(String[] args) throws IOException {
        int port = Integer.parseInt(System.getProperty("APP_PORT", "7070"));
        Path work = Paths.get(System.getProperty("docx.workdir",
                Paths.get(System.getProperty("java.io.tmpdir"), "docx-style-migrator").toString()));
        DocxStyleServer server = new DocxStyleServer(work);
        server.start(port);
    }

    public void start(int port) {
        Javalin app = Javalin.create(config -> {
            config.http.defaultContentType = "application/json";
            config.http.maxRequestSize = 50L * 1024 * 1024; // 50MB
        });

        app.exception(Exception.class, (e, ctx) -> {
            ctx.status(500).json(Map.of("error", e.getMessage()));
        });

        app.get("/", this::sendIndex);
        app.get("/index.html", this::sendIndex);

        app.post("/api/styles/list", ctx -> {
            Path source = saveUpload(ctx, "file", "list-src", Set.of("doc", "docx"));
            try {
                List<DocxStyleService.StyleInfo> styles = service.listStyles(source);
                ctx.json(Map.of("styles", styles));
            } finally {
                Files.deleteIfExists(source);
            }
        });

        app.post("/api/styles/export", ctx -> {
            Path source = saveUpload(ctx, "file", "export-src", Set.of("doc", "docx"));
            Path csv = Files.createTempFile(workDir, "styles-", ".csv");
            try {
                service.exportStyles(source, csv);
                byte[] bytes = Files.readAllBytes(csv);
                ctx.contentType("text/csv; charset=UTF-8");
                ctx.header("Content-Disposition", "attachment; filename=\"styles.csv\"");
                ctx.result(bytes);
            } finally {
                Files.deleteIfExists(source);
                Files.deleteIfExists(csv);
            }
        });

        app.post("/api/styles/migrate", ctx -> {
            // 支持单文件与多文件目标：
            // - 若仅上传一个 targetFile，则行为与之前相同，直接返回处理后的 DOCX
            // - 若上传多个 targetFile，则对每个文件执行迁移，并打包为一个 ZIP 返回
            List<UploadedFile> targetUploads = ctx.uploadedFiles("targetFile");
            if (targetUploads == null || targetUploads.isEmpty()) {
                throw new BadRequestResponse("缺少上传字段: targetFile");
            }

            Path source = saveUploadOrTemplate(ctx, "sourceFile", "src-");
            StyleSelection selection = resolveStyleSelection(ctx, true);
            boolean copyNumbering = parseBoolean(ctx.formParam("copyNumbering"), true);
            boolean includeDependencies = parseBoolean(ctx.formParam("includeDependencies"), true);

            if (targetUploads.size() == 1) {
                // 保持原有单文件行为
                UploadedFile upload = targetUploads.get(0);
                Path target = saveUploadedFile(upload, "dst-", Set.of("doc", "docx"));
                try {
                    Path result;
                    if (selection.allStyles()) {
                        result = service.migrateAll(source, target, copyNumbering);
                    } else {
                        result = service.migrate(source, target, selection.names(), copyNumbering, includeDependencies);
                    }
                    sendDocx(ctx, result, "migrated.docx");
                    if (!result.equals(target)) {
                        deleteQuiet(result);
                    }
                } finally {
                    deleteQuiet(source);
                    deleteQuiet(target);
                }
                return;
            }

            // 多文件批量迁移：返回 ZIP，包内文件名为原始上传文件名
            List<Path> targets = new ArrayList<>();
            List<Path> results = new ArrayList<>();
            List<String> originalNames = new ArrayList<>();
            Path zip = null;
            try {
                for (UploadedFile upload : targetUploads) {
                    if (upload == null || upload.size() == 0) {
                        continue;
                    }
                    Path target = saveUploadedFile(upload, "dst-", Set.of("doc", "docx"));
                    targets.add(target);
                    originalNames.add(safeOriginalName(upload.filename()));
                    Path result;
                    if (selection.allStyles()) {
                        result = service.migrateAll(source, target, copyNumbering);
                    } else {
                        result = service.migrate(source, target, selection.names(), copyNumbering, includeDependencies);
                    }
                    results.add(result);
                }

                if (results.isEmpty()) {
                    throw new BadRequestResponse("未收到有效的目标文件。");
                }

                zip = Files.createTempFile(workDir, "migrated-", ".zip");
                try (OutputStream out = Files.newOutputStream(zip);
                     ZipOutputStream zos = new ZipOutputStream(out)) {
                    Set<String> usedNames = new LinkedHashSet<>();
                    for (int i = 0; i < results.size(); i++) {
                        Path result = results.get(i);
                        String originalName = originalNames.get(i);
                        if (originalName == null || originalName.isBlank()) {
                            originalName = result.getFileName().toString();
                        }
                        String entryName = ensureUniqueName(originalName, usedNames);
                        zos.putNextEntry(new ZipEntry(entryName));
                        try (InputStream in = Files.newInputStream(result)) {
                            in.transferTo(zos);
                        }
                        zos.closeEntry();
                    }
                }

                sendZip(ctx, zip, "migrated.zip");
            } finally {
                deleteQuiet(source);
                for (int i = 0; i < results.size(); i++) {
                    Path result = results.get(i);
                    Path target = i < targets.size() ? targets.get(i) : null;
                    if (result != null && !result.equals(target)) {
                        deleteQuiet(result);
                    }
                }
                for (Path target : targets) {
                    deleteQuiet(target);
                }
                deleteQuiet(zip);
            }
        });

        app.post("/api/styles/format", ctx -> {
            UploadedFile upload = ctx.uploadedFile("file");
            if (upload == null) {
                throw new BadRequestResponse("缺少上传字段: file");
            }
            Path target = saveUploadedFile(upload, "format-", Set.of("doc", "docx"));
            try {
                Path result = service.formatDocument(target);
                sendDocx(ctx, result, buildFormattedFilename(upload.filename()));
                if (!result.equals(target)) {
                    deleteQuiet(result);
                }
            } finally {
                deleteQuiet(target);
            }
        });

        app.post("/api/markdown/convert", ctx -> {
            Path markdown = saveMarkdownInput(ctx, "file", "markdown");
            Path template = saveUploadOrTemplate(ctx, "templateFile", "md-template-");
            String title = ctx.formParam("title");
            String sourceName = resolveMarkdownSourceName(ctx.uploadedFile("file"), title);
            try {
                Path result = service.convertMarkdown(markdown, template, title);
                sendDocx(ctx, result, buildMarkdownFilename(sourceName));
                deleteQuiet(result);
            } finally {
                deleteQuiet(markdown);
                deleteQuiet(template);
            }
        });

        app.post("/api/styles/clean", ctx -> {
            Path target = saveUpload(ctx, "file", "clean-", Set.of("doc", "docx"));
            StyleSelection selection = resolveStyleSelection(ctx, false);
            if (selection.allStyles()) {
                throw new BadRequestResponse("清理操作不支持通配符 *");
            }
            try {
                DocxStyleService.CleanResult result = service.cleanStyles(target, selection.names());
                ctx.header("X-Removed-Count", Integer.toString(result.removed()));
                sendDocx(ctx, result.file(), "cleaned.docx");
                if (!result.file().equals(target)) {
                    deleteQuiet(result.file());
                }
            } finally {
                deleteQuiet(target);
            }
        });

        app.start(port);
    }

    private void sendDocx(Context ctx, Path file, String filename) throws IOException {
        byte[] bytes = Files.readAllBytes(file);
        ctx.contentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        ctx.header("Content-Disposition", buildAttachmentDisposition(filename));
        ctx.result(bytes);
    }

    private void sendIndex(Context ctx) throws IOException {
        ctx.contentType("text/html; charset=UTF-8");
        ctx.result(readClasspathResource("/public/index.html"));
    }

    private void sendZip(Context ctx, Path file, String filename) throws IOException {
        byte[] bytes = Files.readAllBytes(file);
        ctx.contentType("application/zip");
        ctx.header("Content-Disposition", buildAttachmentDisposition(filename));
        ctx.result(bytes);
    }

    private String buildAttachmentDisposition(String filename) {
        String safeFilename = safeOriginalName(filename);
        String asciiFilename = safeFilename.replaceAll("[^\\x20-\\x7E]", "_").replace("\"", "");
        if (asciiFilename.isBlank()) {
            asciiFilename = "download.docx";
        }
        String encodedFilename = URLEncoder.encode(safeFilename, StandardCharsets.UTF_8)
                .replace("+", "%20");
        return "attachment; filename=\"" + asciiFilename + "\"; filename*=UTF-8''" + encodedFilename;
    }

    private String ensureUniqueName(String name, Set<String> used) {
        String base = name;
        String ext = "";
        int dot = name.lastIndexOf('.');
        if (dot > 0 && dot < name.length() - 1) {
            base = name.substring(0, dot);
            ext = name.substring(dot);
        }
        String candidate = name;
        int index = 1;
        while (used.contains(candidate)) {
            candidate = base + "(" + index + ")" + ext;
            index++;
        }
        used.add(candidate);
        return candidate;
    }

    private Path saveUploadOrTemplate(Context ctx, String field, String prefix) throws IOException {
        UploadedFile upload = ctx.uploadedFile(field);
        if (upload == null || upload.size() == 0) {
            return extractTemplate(prefix);
        }
        return saveUploadedFile(upload, prefix, Set.of("doc", "docx"));
    }

    private Path saveUpload(Context ctx, String field, String prefix, Set<String> allowedExtensions) throws IOException {
        UploadedFile upload = ctx.uploadedFile(field);
        if (upload == null) {
            throw new BadRequestResponse("缺少上传字段: " + field);
        }
        return saveUploadedFile(upload, prefix, allowedExtensions);
    }

    private Path saveMarkdownInput(Context ctx, String fileField, String textField) throws IOException {
        UploadedFile upload = ctx.uploadedFile(fileField);
        if (upload != null && upload.size() > 0) {
            return saveTextFile(upload, "markdown-", Set.of("md", "markdown", "txt"));
        }
        String markdown = ctx.formParam(textField);
        if (markdown == null || markdown.isBlank()) {
            throw new BadRequestResponse("缺少 Markdown 内容，请上传文件或填写文本。");
        }
        Path temp = Files.createTempFile(workDir, "markdown-", ".md");
        Files.writeString(temp, markdown, StandardCharsets.UTF_8);
        return temp;
    }

    private Path saveUploadedFile(UploadedFile upload, String prefix, Set<String> allowedExtensions) throws IOException {
        String originalName = upload.filename();
        Path temp = Files.createTempFile(workDir, prefix, ".tmp");
        try (var in = upload.content()) {
            Files.copy(in, temp, StandardCopyOption.REPLACE_EXISTING);
        }
        String detectedExtension;
        try {
            detectedExtension = detectDocExtension(temp, originalName);
        } catch (IOException | RuntimeException ex) {
            deleteQuiet(temp);
            throw ex;
        }
        if (allowedExtensions != null && !allowedExtensions.isEmpty() && !allowedExtensions.contains(detectedExtension)) {
            deleteQuiet(temp);
            throw new BadRequestResponse("文件格式不被允许: " + safeOriginalName(originalName));
        }
        return ensureExtension(temp, detectedExtension);
    }

    private Path saveTextFile(UploadedFile upload, String prefix, Set<String> allowedExtensions) throws IOException {
        String originalName = safeOriginalName(upload.filename());
        String extension = FilenameUtils.getExtension(originalName).toLowerCase(Locale.ROOT);
        if (extension.isBlank()) {
            extension = "md";
        }
        if (allowedExtensions != null && !allowedExtensions.isEmpty() && !allowedExtensions.contains(extension)) {
            throw new BadRequestResponse("文件格式不被允许: " + originalName);
        }
        Path temp = Files.createTempFile(workDir, prefix, "." + extension);
        try (var in = upload.content()) {
            Files.copy(in, temp, StandardCopyOption.REPLACE_EXISTING);
        }
        return temp;
    }

    private String detectDocExtension(Path file, String originalName) throws IOException {
        byte[] header = new byte[8];
        int read;
        try (var in = Files.newInputStream(file)) {
            read = in.read(header);
        }
        if (read >= 4) {
            if ((header[0] & 0xFF) == 0x50 && (header[1] & 0xFF) == 0x4B && (header[2] & 0xFF) == 0x03 && (header[3] & 0xFF) == 0x04) {
                return "docx";
            }
        }
        if (read >= 8) {
            int[] ole = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
            boolean match = true;
            for (int i = 0; i < ole.length; i++) {
                if ((header[i] & 0xFF) != ole[i]) {
                    match = false;
                    break;
                }
            }
            if (match) {
                return "doc";
            }
        }
        String extensionHint = originalName == null ? "" : FilenameUtils.getExtension(originalName).toLowerCase(Locale.ROOT);
        if ("doc".equals(extensionHint) || "docx".equals(extensionHint)) {
            return extensionHint;
        }
        throw new BadRequestResponse("无法识别的文档格式: " + safeOriginalName(originalName));
    }

    private Path ensureExtension(Path file, String extension) throws IOException {
        String filename = file.getFileName().toString();
        int dot = filename.lastIndexOf('.');
        String base = dot >= 0 ? filename.substring(0, dot) : filename;
        Path target = file.resolveSibling(base + "." + extension);
        return Files.move(file, target, StandardCopyOption.REPLACE_EXISTING);
    }

    private String safeOriginalName(String name) {
        if (name == null || name.isBlank()) {
            return "(未知文件)";
        }
        return name;
    }

    private String buildFormattedFilename(String originalName) {
        String safeName = safeOriginalName(originalName);
        String baseName = FilenameUtils.getBaseName(safeName);
        if (baseName == null || baseName.isBlank()) {
            baseName = "document";
        }
        return baseName + "_已格式化.docx";
    }

    private String buildMarkdownFilename(String originalName) {
        String safeName = safeOriginalName(originalName);
        String baseName = FilenameUtils.getBaseName(safeName);
        if (baseName == null || baseName.isBlank()) {
            baseName = "markdown";
        }
        return baseName + "_已转换.docx";
    }

    private String resolveMarkdownSourceName(UploadedFile upload, String title) {
        if (upload != null && upload.filename() != null && !upload.filename().isBlank()) {
            return upload.filename();
        }
        if (title != null && !title.isBlank()) {
            return title.trim() + ".md";
        }
        return "markdown.md";
    }

    private Path extractTemplate(String prefix) throws IOException {
        String resourcePath = "/template/template.docx";
        try (var stream = DocxStyleServer.class.getResourceAsStream(resourcePath)) {
            if (stream == null) {
                throw new IllegalStateException("找不到内置模板: " + resourcePath);
            }
            Path temp = Files.createTempFile(workDir, prefix, ".docx");
            Files.copy(stream, temp, StandardCopyOption.REPLACE_EXISTING);
            return temp;
        }
    }

    private String readClasspathResource(String resourcePath) throws IOException {
        try (InputStream stream = DocxStyleServer.class.getResourceAsStream(resourcePath)) {
            if (stream == null) {
                throw new IllegalStateException("找不到内置资源: " + resourcePath);
            }
            return new String(stream.readAllBytes(), StandardCharsets.UTF_8);
        }
    }

    private StyleSelection resolveStyleSelection(Context ctx, boolean allowWildcard) throws IOException {
        Set<String> styles = new LinkedHashSet<>();
        boolean wildcard = false;
        wildcard |= addTokens(styles, ctx.formParam("styles"));
        UploadedFile styleFile = ctx.uploadedFile("stylesFile");
        if (styleFile != null) {
            try (BufferedReader reader = new BufferedReader(
                    new InputStreamReader(styleFile.content(), StandardCharsets.UTF_8))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    wildcard |= addTokens(styles, line);
                }
            }
        }
        if (wildcard) {
            if (!allowWildcard) {
                throw new BadRequestResponse("当前操作不支持通配符 *");
            }
            return new StyleSelection(Set.of(), true);
        }
        if (styles.isEmpty()) {
            throw new BadRequestResponse("至少需要指定一个样式名称或 ID。");
        }
        return new StyleSelection(styles, false);
    }

    private boolean addTokens(Set<String> bucket, String source) {
        if (source == null) {
            return false;
        }
        boolean wildcard = false;
        for (String token : source.split(",")) {
            String value = token.trim();
            if (value.isEmpty()) {
                continue;
            }
            if ("*".equals(value)) {
                wildcard = true;
            } else {
                bucket.add(value);
            }
        }
        return wildcard;
    }

    private boolean parseBoolean(String value, boolean defaultValue) {
        if (value == null || value.isBlank()) {
            return defaultValue;
        }
        if ("true".equalsIgnoreCase(value) || "y".equalsIgnoreCase(value) || "yes".equalsIgnoreCase(value)) {
            return true;
        }
        if ("false".equalsIgnoreCase(value) || "n".equalsIgnoreCase(value) || "no".equalsIgnoreCase(value)) {
            return false;
        }
        return defaultValue;
    }

    private void deleteQuiet(Path path) {
        if (path == null) {
            return;
        }
        try {
            Files.deleteIfExists(path);
        } catch (IOException ignore) {
        }
    }

    private record StyleSelection(Set<String> names, boolean allStyles) {
    }
}
