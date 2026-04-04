package io.github.connellite.excel;

import io.github.connellite.excel.exception.ExcelSheetRowException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.NoSuchElementException;

/**
 * Forward-only iterator over sheet rows with {@link Map} values taken from POI cell types (strings, numbers as
 * {@link Double} for plain numbers, booleans, formula results). Date-formatted numeric cells are returned as formatted strings.
 * Implements {@link Iterable} and {@link AutoCloseable} with the base class contract; {@link #close()} closes the workbook.
 *
 * @see ExcelRowStringIterator
 * @see ExcelRowStream#stream(Workbook, String)
 */
public class ExcelRowIterator extends AbstractExcelSheetRowIterator<Object> {

    /**
     * Creates an iterator over the given sheet. Closing this iterator closes {@code workbook}.
     */
    public ExcelRowIterator(Workbook workbook, String sheetName) {
        super(workbook, sheetName);
    }

    @Override
    public Map<String, Object> next() {
        if (!hasNext()) {
            throw new NoSuchElementException();
        }

        Row row = rowIterator.next();
        Map<String, Object> rowData = new LinkedHashMap<>();
        try {
            for (int i = 0; i < columnNames.size(); i++) {
                Cell cell = row.getCell(i);
                rowData.put(columnNames.get(i), getCellObjectValue(cell));
            }
        } catch (Exception e) {
            throw new ExcelSheetRowException(e);
        }

        return rowData;
    }
}
