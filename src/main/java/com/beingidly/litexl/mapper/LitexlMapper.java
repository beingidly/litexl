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

/**
 * Maps Java objects to and from Excel workbooks using annotations.
 */
public final class LitexlMapper {

    private static final Set<Class<?>> registeredClasses = ConcurrentHashMap.newKeySet();

    private final MapperConfig config;

    private LitexlMapper(MapperConfig config) {
        this.config = config;
    }

    /**
     * Reads an Excel file and maps it to the given type.
     *
     * @param <T> the target type
     * @param path the file path
     * @param type the target class
     * @return the mapped object
     */
    public static <T> T read(Path path, Class<T> type) {
        return new LitexlMapper(MapperConfig.defaults()).doRead(path, null, type);
    }

    /**
     * Reads an encrypted Excel file and maps it to the given type.
     *
     * @param <T> the target type
     * @param path the file path
     * @param password the decryption password
     * @param type the target class
     * @return the mapped object
     */
    public static <T> T read(Path path, String password, Class<T> type) {
        return new LitexlMapper(MapperConfig.defaults()).doRead(path, password, type);
    }

    /**
     * Writes the given object to an Excel file.
     *
     * @param <T> the source type
     * @param object the object to write
     * @param path the output file path
     */
    public static <T> void write(T object, Path path) {
        new LitexlMapper(MapperConfig.defaults()).doWrite(object, path, null);
    }

    /**
     * Writes the given object to an encrypted Excel file.
     *
     * @param <T> the source type
     * @param object the object to write
     * @param path the output file path
     * @param options the encryption options
     */
    public static <T> void write(T object, Path path, EncryptionOptions options) {
        new LitexlMapper(MapperConfig.defaults()).doWrite(object, path, options);
    }

    /**
     * Reads an Excel file and maps it to the given type using this mapper's config.
     *
     * @param <T> the target type
     * @param path the file path
     * @param type the target class
     * @return the mapped object
     */
    public <T> T readFile(Path path, Class<T> type) {
        return doRead(path, null, type);
    }

    /**
     * Reads an encrypted Excel file and maps it to the given type using this mapper's config.
     *
     * @param <T> the target type
     * @param path the file path
     * @param password the decryption password
     * @param type the target class
     * @return the mapped object
     */
    public <T> T readFile(Path path, String password, Class<T> type) {
        return doRead(path, password, type);
    }

    /**
     * Writes the given object to an Excel file using this mapper's config.
     *
     * @param <T> the source type
     * @param object the object to write
     * @param path the output file path
     */
    public <T> void writeFile(T object, Path path) {
        doWrite(object, path, null);
    }

    /**
     * Writes the given object to an encrypted Excel file using this mapper's config.
     *
     * @param <T> the source type
     * @param object the object to write
     * @param path the output file path
     * @param options the encryption options
     */
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

    /**
     * Creates a new mapper builder.
     *
     * @return a new builder
     */
    public static Builder builder() {
        return new Builder();
    }

    /** Builder for configuring a LitexlMapper instance. */
    public static final class Builder {
        private String dateFormat = "yyyy-MM-dd HH:mm:ss";
        private NullStrategy nullStrategy = NullStrategy.SKIP;

        private Builder() {}

        /**
         * Sets the date format pattern.
         *
         * @param format the date format pattern
         * @return this builder
         */
        public Builder dateFormat(String format) {
            this.dateFormat = format;
            return this;
        }

        /**
         * Sets the null handling strategy.
         *
         * @param strategy the null strategy
         * @return this builder
         */
        public Builder nullStrategy(NullStrategy strategy) {
            this.nullStrategy = strategy;
            return this;
        }

        /**
         * Builds a new LitexlMapper with the configured options.
         *
         * @return a new mapper
         */
        public LitexlMapper build() {
            return new LitexlMapper(new MapperConfig(dateFormat, nullStrategy));
        }
    }

    /**
     * Registers classes for GraalVM native image support.
     *
     * @param classes the classes to register
     */
    public static void register(Class<?>... classes) {
        registeredClasses.addAll(Arrays.asList(classes));
    }

    /**
     * Returns all registered classes.
     *
     * @return an unmodifiable set of registered classes
     */
    public static Set<Class<?>> getRegisteredClasses() {
        return Set.copyOf(registeredClasses);
    }
}
