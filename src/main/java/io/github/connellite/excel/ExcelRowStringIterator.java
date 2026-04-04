package io.github.connellite.excel;

import io.github.connellite.excel.exception.ExcelSheetRowException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.math.BigDecimal;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.NoSuchElementException;

/**
 * Forward-only iterator over sheet rows with all cell values exposed as {@link String}s. Numbers use a plain
 * string without scientific notation (trailing zeros stripped); date-formatted cells use the data formatter output.
 * Formula cells expose the formula expression as text. Implements {@link Iterable} and {@link AutoCloseable} per the
 * base class; {@link #close()} closes the workbook.
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
                rowData.put(columnNames.get(i), getCellValue(cell));
            }
        } catch (Exception e) {
            throw new ExcelSheetRowException(e);
        }

        return rowData;
    }

    private String getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> DateUtil.isCellDateFormatted(cell)
                    ? formatter.formatCellValue(cell)
                    : BigDecimal.valueOf(cell.getNumericCellValue()).stripTrailingZeros().toPlainString();
            case BOOLEAN -> Boolean.toString(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            case BLANK -> null;
            default -> formatter.formatCellValue(cell);
        };
    }
}
