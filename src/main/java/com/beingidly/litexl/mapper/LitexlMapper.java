package com.beingidly.litexl.mapper;

import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.crypto.EncryptionOptions;
import com.beingidly.litexl.mapper.internal.MapperConfig;
import com.beingidly.litexl.mapper.internal.WorkbookReader;
import com.beingidly.litexl.mapper.internal.WorkbookWriter;
import org.jspecify.annotations.Nullable;

import java.nio.file.Path;
import java.util.Arrays;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

public final class LitexlMapper {

    private static final Set<Class<?>> registeredClasses = ConcurrentHashMap.newKeySet();

    private final MapperConfig config;

    private LitexlMapper(MapperConfig config) {
        this.config = config;
    }

    // Static methods (default config)

    public static <T> T read(Path path, Class<T> type) {
        return new LitexlMapper(MapperConfig.defaults()).doRead(path, null, type);
    }

    public static <T> T read(Path path, String password, Class<T> type) {
        return new LitexlMapper(MapperConfig.defaults()).doRead(path, password, type);
    }

    public static <T> void write(T object, Path path) {
        new LitexlMapper(MapperConfig.defaults()).doWrite(object, path, null);
    }

    public static <T> void write(T object, Path path, EncryptionOptions options) {
        new LitexlMapper(MapperConfig.defaults()).doWrite(object, path, options);
    }

    // Instance methods (custom config) - these are for builder-created instances

    public <T> T readFile(Path path, Class<T> type) {
        return doRead(path, null, type);
    }

    public <T> T readFile(Path path, String password, Class<T> type) {
        return doRead(path, password, type);
    }

    public <T> void writeFile(T object, Path path) {
        doWrite(object, path, null);
    }

    public <T> void writeFile(T object, Path path, EncryptionOptions options) {
        doWrite(object, path, options);
    }

    private <T> T doRead(Path path, @Nullable String password, Class<T> type) {
        try (Workbook workbook = password != null
                ? Workbook.open(path, password)
                : Workbook.open(path)) {
            var reader = new WorkbookReader(config);
            return reader.read(workbook, type);
        }
    }

    private <T> void doWrite(T object, Path path, @Nullable EncryptionOptions options) {
        try (Workbook workbook = Workbook.create()) {
            var writer = new WorkbookWriter(config);
            writer.write(workbook, object);

            if (options != null) {
                workbook.save(path, options);
            } else {
                workbook.save(path);
            }
        }
    }

    // Builder

    public static Builder builder() {
        return new Builder();
    }

    public static final class Builder {
        private String dateFormat = "yyyy-MM-dd HH:mm:ss";
        private NullStrategy nullStrategy = NullStrategy.SKIP;

        private Builder() {}

        public Builder dateFormat(String format) {
            this.dateFormat = format;
            return this;
        }

        public Builder nullStrategy(NullStrategy strategy) {
            this.nullStrategy = strategy;
            return this;
        }

        public LitexlMapper build() {
            return new LitexlMapper(new MapperConfig(dateFormat, nullStrategy));
        }
    }

    // GraalVM support

    public static void register(Class<?>... classes) {
        registeredClasses.addAll(Arrays.asList(classes));
    }

    public static Set<Class<?>> getRegisteredClasses() {
        return Set.copyOf(registeredClasses);
    }
}
