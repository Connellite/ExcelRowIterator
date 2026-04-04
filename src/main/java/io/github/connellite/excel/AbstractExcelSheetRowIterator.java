package io.github.connellite.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Forward-only iterator over data rows of an Apache POI {@link Sheet}. The first row is treated as a header;
 * column names are derived from that row (blank headers become {@code Column_0}, {@code Column_1}, …).
 * Implements {@link Iterable} so you can use {@code for (Map<String, V> row : it)} together with
 * try-with-resources. {@link #iterator()} returns {@code this}; the cursor is single-pass, so a second
 * enhanced for-loop on the same instance does not replay rows from the start.
 * <p>
 * {@link #close()} closes the {@link Workbook} passed to the constructor. Do not close the same workbook
 * again afterward (for example avoid wrapping the workbook in a second try-with-resources if the iterator
 * already owns shutdown).
 *
 * @param <V> map value type produced by subclasses
 * @see ExcelRowIterator
 * @see ExcelRowStringIterator
 */
public abstract class AbstractExcelSheetRowIterator<V> implements Iterator<Map<String, V>>, Iterable<Map<String, V>>, AutoCloseable {

    private final Workbook workbook;
    protected final List<String> columnNames;
    protected final Iterator<Row> rowIterator;
    protected final DataFormatter formatter = new DataFormatter();

    /**
     * @param workbook  workbook that contains the sheet; closed when {@link #close()} is called
     * @param sheetName name of the sheet to read
     * @throws IllegalArgumentException if no sheet exists with the given name
     */
    protected AbstractExcelSheetRowIterator(Workbook workbook, String sheetName) {
        this.workbook = workbook;
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet not found: " + sheetName);
        }

        rowIterator = sheet.iterator();
        if (rowIterator.hasNext()) {
            Row headerRow = rowIterator.next();
            int columnCount = Math.max(headerRow.getLastCellNum(), 0);
            List<String> names = new ArrayList<>(columnCount);
            for (int i = 0; i < columnCount; i++) {
                Cell cell = headerRow.getCell(i);
                if (cell == null) {
                    names.add("Column_" + i);
                    continue;
                }
                String value = formatter.formatCellValue(cell);
                names.add(isBlank(value) ? "Column_" + i : value);
            }
            this.columnNames = names;
        } else {
            this.columnNames = Collections.emptyList();
        }
    }

    private static boolean isBlank(String value) {
        return value == null || value.isBlank();
    }

    /**
     * {@inheritDoc}
     * <p>
     * Returns {@code this} as the only iterator; one traversal consumes the sheet rows after the header.
     */
    @Override
    public Iterator<Map<String, V>> iterator() {
        return this;
    }

    @Override
    public boolean hasNext() {
        return rowIterator.hasNext();
    }

    @Override
    public abstract Map<String, V> next();

    /**
     * Column names from the first row of the sheet, in order.
     */
    public List<String> getColumnNames() {
        return Collections.unmodifiableList(columnNames);
    }

    /**
     * Closes the workbook from construction.
     */
    @Override
    public void close() throws Exception {
        try {
            if (workbook != null) {
                workbook.close();
            }
        } catch (Exception ignore) {
        }
    }
}
