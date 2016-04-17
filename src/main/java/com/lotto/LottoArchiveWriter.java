package com.lotto;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.StringJoiner;

/**
 * Created by Alex on 12.04.2016.
 */
public class LottoArchiveWriter {
    public static void main(String[] args) throws Exception {
        FileInputStream file = new FileInputStream(new File("LOTTO6aus49_1.xlsx"));
        BufferedWriter out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("lottoNumbers.txt"), "UTF-8"));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        for (int year = 2001; year < 2016; year++) {
            XSSFSheet sheet = workbook.getSheet(String.valueOf(year));
            int rowCount = sheet.getLastRowNum();
            for (int i = 7; i < rowCount; i++) {
                Row row = sheet.getRow(i);
                Cell weekDayCell = row.getCell(2);
                String weekDay = getCellValue(weekDayCell);
                if (weekDay.equals("MI")) {
                    StringJoiner stringJoiner = new StringJoiner(" ");
                    Cell dayCell = row.getCell(1);
                    String day = getCellValue(dayCell);
                    stringJoiner.add(day);

                    for (int l = 3; l < 9; l++) {
                        Cell lottoNumberCell = row.getCell(l);
                        String lottoNumber = getCellValue(lottoNumberCell);
                        stringJoiner.add(lottoNumber);
                    }
                    out.write(stringJoiner.toString());
                    out.newLine();
                    System.out.println(stringJoiner.toString());
                }
            }
        }
        out.close();
    }

    public static String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case Cell.CELL_TYPE_NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
                    return dateFormat.format(date);
                } else {
                    return String.valueOf(Double.valueOf(cell.getNumericCellValue()).intValue());
                }
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
//            case Cell.
            default:
                return cell.getStringCellValue();
        }
    }
}
