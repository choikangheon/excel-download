package com.kangheon.excelDownload;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelDownloadTest {
    @Test
    public void testHSSFPerformance() { // 20만건을 메모리에 고대로 올려서 작업
        long startTime = System.currentTimeMillis();
        Runtime runtime = Runtime.getRuntime();
        long initialMemory = runtime.totalMemory() - runtime.freeMemory();

        Workbook hssfWorkbook = new HSSFWorkbook();
        int totalRowCount = 200000; // 20만 건의 데이터
        int batchSize = 65000; // 65,000 개 단위로 시트를 나눕니다.

        for (int i = 0; i < totalRowCount; i++) {
            if (i % batchSize == 0) {
                // batchSize만큼의 데이터가 작성되면 새로운 시트를 생성합니다.
                createSheet(hssfWorkbook, i / batchSize);
            }

            // 데이터 생성 코드
            Sheet sheet = hssfWorkbook.getSheetAt(i / batchSize);
            Row row = sheet.createRow(i % batchSize);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue("Data" + i + j);
            }
        }

        long endTime = System.currentTimeMillis();
        System.out.println("HSSF time taken: " + (endTime - startTime) + " ms");

        long finalMemory = runtime.totalMemory() - runtime.freeMemory();
        long memoryUsed = finalMemory - initialMemory;
        System.out.println("HSSF memory used: " + memoryUsed + " bytes");
    }

    private void createSheet(Workbook workbook, int sheetIndex) {
        workbook.createSheet("TestSheet" + sheetIndex);
    }

    @Test
    public void testSXSSFPerformance() {
        try {
            long startTime = System.currentTimeMillis();

            Runtime runtime = Runtime.getRuntime();
            long initialMemory = runtime.totalMemory() - runtime.freeMemory();

            Workbook sxssfWorkbook = new SXSSFWorkbook(5);
            Sheet sheet = sxssfWorkbook.createSheet("TestSheet");

            for (int i = 0; i < 20; i++) { // 20만 건의 데이터 생성
                Row row = sheet.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue("Data" + i);
            }


            FileOutputStream fos = new FileOutputStream(new File("/Users/kangheon/Desktop/test.xlsx"));
            sxssfWorkbook.write(fos);
            fos.close();

            long endTime = System.currentTimeMillis();
            System.out.println("SXSSF time taken: " + (endTime - startTime) + " ms");

            long finalMemory = runtime.totalMemory() - runtime.freeMemory();
            long memoryUsed = finalMemory - initialMemory;
            System.out.println("SXSSF memory used: " + memoryUsed + " bytes");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
