package io.github.connellite.excel;

import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Builds a sequential {@link Stream} over sheet rows. Close the stream (for example with try-with-resources)
 * to run {@link AutoCloseable#close()} on the underlying iterator, which closes the {@link Workbook}.
 */
public final class ExcelRowStream {

    private ExcelRowStream() {
    }

    /**
     * Opens a forward-only, ordered stream of rows as {@code Map<String, Object>} (see {@link ExcelRowIterator}).
     *
     * @param workbook  workbook to read from; closed when the stream is closed
     * @param sheetName sheet name
     * @return a non-parallel stream
     */
    public static Stream<Map<String, Object>> stream(Workbook workbook, String sheetName) {
        ExcelRowIterator it = new ExcelRowIterator(workbook, sheetName);
        return StreamSupport.stream(
                        Spliterators.spliteratorUnknownSize(it, Spliterator.ORDERED | Spliterator.NONNULL),
                        false)
                .onClose(() -> {
                    try {
                        it.close();
                    } catch (Exception e) {
                        throw new RuntimeException(e);
                    }
                });
    }

    /**
     * Opens a forward-only, ordered stream of rows as {@code Map<String, String>} (see {@link ExcelRowStringIterator}).
     *
     * @param workbook  workbook to read from; closed when the stream is closed
     * @param sheetName sheet name
     * @return a non-parallel stream
     */
    public static Stream<Map<String, String>> streamStrings(Workbook workbook, String sheetName) {
        ExcelRowStringIterator it = new ExcelRowStringIterator(workbook, sheetName);
        return StreamSupport.stream(
                        Spliterators.spliteratorUnknownSize(it, Spliterator.ORDERED | Spliterator.NONNULL),
                        false)
                .onClose(() -> {
                    try {
                        it.close();
                    } catch (Exception e) {
                        throw new RuntimeException(e);
                    }
                });
    }
}
