package com.example.docx;

import io.javalin.Javalin;
import io.javalin.http.BadRequestResponse;
import io.javalin.http.Context;
import io.javalin.http.UploadedFile;
import org.apache.commons.io.FilenameUtils;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

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
            config.staticFiles.add("/public");
        });

        app.exception(Exception.class, (e, ctx) -> {
            ctx.status(500).json(Map.of("error", e.getMessage()));
        });

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
            Path source = saveUploadOrTemplate(ctx, "sourceFile", "src-");
            Path target = saveUpload(ctx, "targetFile", "dst-", Set.of("doc", "docx"));
            StyleSelection selection = resolveStyleSelection(ctx, true);
            boolean copyNumbering = parseBoolean(ctx.formParam("copyNumbering"), true);
            boolean includeDependencies = parseBoolean(ctx.formParam("includeDependencies"), true);
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
        ctx.header("Content-Disposition", "attachment; filename=\"" + filename + "\"");
        ctx.result(bytes);
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
