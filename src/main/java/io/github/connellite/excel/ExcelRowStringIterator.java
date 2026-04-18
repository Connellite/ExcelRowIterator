package io.github.connellite.excel;

import io.github.connellite.excel.exception.ExcelSheetRowException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.math.BigDecimal;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Objects;

/**
 * Forward-only iterator over sheet rows with all cell values exposed as {@link String}s. Numbers use a plain
 * string without scientific notation (trailing zeros stripped); date-formatted cells use the data formatter output.
 * Formula cells use the cached result (same basis as {@link ExcelRowIterator}), not the formula text.
 * Implements {@link Iterable} and {@link AutoCloseable} per the base class; {@link #close()} closes the workbook.
 *
 * @see ExcelRowIterator
 * @see ExcelRowStream#streamStrings(Workbook, String)
 */
public class ExcelRowStringIterator extends AbstractExcelSheetRowIterator<String> {

    /**
     * Creates a string-valued iterator over the given sheet. Closing this iterator closes {@code workbook}.
     */
    public ExcelRowStringIterator(Workbook workbook, String sheetName) {
        super(workbook, sheetName);
    }

    @Override
    public Map<String, String> next() {
        if (!hasNext()) {
            throw new NoSuchElementException();
        }

        Row row = rowIterator.next();
        Map<String, String> rowData = new LinkedHashMap<>();
        try {
            for (int i = 0; i < columnNames.size(); i++) {
                Cell cell = row.getCell(i);
                rowData.put(columnNames.get(i), objectToPlainString(getCellObjectValue(cell)));
            }
        } catch (Exception e) {
            throw new ExcelSheetRowException(e);
        }

        return Collections.unmodifiableMap(rowData);
    }

    private static String objectToPlainString(Object value) {
        if (value instanceof Number n) {
            return new BigDecimal(n.toString()).stripTrailingZeros().toPlainString();
        }
        return Objects.toString(value, null);
    }
}
