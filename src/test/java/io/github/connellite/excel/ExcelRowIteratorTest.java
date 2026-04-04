package io.github.connellite.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelRowIteratorTest {

    @Test
    void typedIterator_mapsColumnsAndTypes() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sh = wb.createSheet("Data");
        Row h = sh.createRow(0);
        h.createCell(0).setCellValue("name");
        h.createCell(1).setCellValue("amount");
        Row r1 = sh.createRow(1);
        r1.createCell(0).setCellValue("a");
        r1.createCell(1).setCellValue(42.5);

        try (ExcelRowIterator it = new ExcelRowIterator(wb, "Data")) {
            assertEquals(List.of("name", "amount"), it.getColumnNames());
            assertTrue(it.hasNext());
            Map<String, Object> row = it.next();
            assertEquals("a", row.get("name"));
            assertInstanceOf(BigDecimal.class, row.get("amount"));
            assertEquals(new BigDecimal("42.5"), row.get("amount"));
            assertFalse(it.hasNext());
        }
    }

    @Test
    void stringIterator_stripsNumericTrailingZeros() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sh = wb.createSheet("S");
        Row h = sh.createRow(0);
        h.createCell(0).setCellValue("n");
        Row r = sh.createRow(1);
        r.createCell(0).setCellValue(10.0);

        try (ExcelRowStringIterator it = new ExcelRowStringIterator(wb, "S")) {
            Map<String, String> row = it.next();
            assertEquals("10", row.get("n"));
        }
    }

    @Test
    void blankHeader_usesSyntheticName() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sh = wb.createSheet("S");
        Row h = sh.createRow(0);
        h.createCell(0).setCellValue("");
        h.createCell(1).setCellValue("b");
        Row r = sh.createRow(1);
        r.createCell(0).setCellValue("x");
        r.createCell(1).setCellValue("y");

        try (ExcelRowStringIterator it = new ExcelRowStringIterator(wb, "S")) {
            Map<String, String> row = it.next();
            assertEquals("x", row.get("Column_0"));
            assertEquals("y", row.get("b"));
        }
    }

    @Test
    void emptySheet_noDataRows() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        wb.createSheet("Empty");
        try (ExcelRowIterator it = new ExcelRowIterator(wb, "Empty")) {
            assertTrue(it.getColumnNames().isEmpty());
            assertFalse(it.hasNext());
        }
    }

    @Test
    void missingSheet_throws() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("Only");
            assertThrows(IllegalArgumentException.class, () -> new ExcelRowIterator(wb, "Nope"));
        }
    }

    @Test
    void nextPastEnd_throwsNoSuchElement() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sh = wb.createSheet("S");
        sh.createRow(0).createCell(0).setCellValue("c");
        try (ExcelRowIterator it = new ExcelRowIterator(wb, "S")) {
            assertThrows(NoSuchElementException.class, it::next);
        }
    }

    @Test
    void stream_collectsRows() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sh = wb.createSheet("D");
        sh.createRow(0).createCell(0).setCellValue("k");
        sh.createRow(1).createCell(0).setCellValue("v1");
        sh.createRow(2).createCell(0).setCellValue("v2");

        try (Stream<Map<String, Object>> s = ExcelRowStream.stream(wb, "D")) {
            List<Map<String, Object>> rows = s.toList();
            assertEquals(2, rows.size());
            assertEquals("v1", rows.get(0).get("k"));
            assertEquals("v2", rows.get(1).get("k"));
        }
    }

    @Test
    void streamStrings_collectsRows() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sh = wb.createSheet("D");
        sh.createRow(0).createCell(0).setCellValue("k");
        sh.createRow(1).createCell(0).setCellValue("1");

        try (Stream<Map<String, String>> s = ExcelRowStream.streamStrings(wb, "D")) {
            List<Map<String, String>> rows = new ArrayList<>(s.toList());
            assertEquals(1, rows.size());
            assertEquals("1", rows.get(0).get("k"));
        }
    }

    @Test
    void booleanCell_typedAndString() throws Exception {
        XSSFWorkbook wb1 = new XSSFWorkbook();
        Sheet sh1 = wb1.createSheet("B");
        sh1.createRow(0).createCell(0).setCellValue("flag");
        sh1.createRow(1).createCell(0).setCellValue(true);
        try (ExcelRowIterator typed = new ExcelRowIterator(wb1, "B")) {
            assertEquals(Boolean.TRUE, typed.next().get("flag"));
        }

        XSSFWorkbook wb2 = new XSSFWorkbook();
        Sheet sh2 = wb2.createSheet("B2");
        sh2.createRow(0).createCell(0).setCellValue("flag");
        sh2.createRow(1).createCell(0).setCellValue(false);
        try (ExcelRowStringIterator str = new ExcelRowStringIterator(wb2, "B2")) {
            assertEquals("false", str.next().get("flag"));
        }
    }

    @Test
    void nullCell_yieldsNullValue() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sh = wb.createSheet("N");
        sh.createRow(0).createCell(0).setCellValue("a");
        sh.createRow(1).createCell(0); // blank

        try (ExcelRowStringIterator it = new ExcelRowStringIterator(wb, "N")) {
            assertNull(it.next().get("a"));
        }
    }
}
