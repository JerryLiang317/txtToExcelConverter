package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class TxtToExcelMethod {

    public static void txtToExcel(String readFolder, String outExcelFile) throws IOException {
        File[] txtFiles = new File(readFolder).listFiles();

        if (txtFiles != null) {
            for (File txtFile : txtFiles) {
                if (txtFile.isFile() && txtFile.getName().endsWith(".txt")) {
                    List<String[]> data = readTxtFile(txtFile);

                    Workbook workbook = new XSSFWorkbook();
                    Sheet sheet = workbook.createSheet("Sheet1");

                    int rowNum = 0;
                    for (String[] rowData : data) {
                        Row row = sheet.createRow(rowNum++);
                        int cellNum = 0;
                        for (String cellData : rowData) {
                            Cell cell = row.createCell(cellNum++);
                            try {
                                double numericValue = Double.parseDouble(cellData);
                                cell.setCellValue(numericValue);
                            } catch (NumberFormatException e) {
                                cell.setCellValue(cellData);
                            }
                        }
                    }

                    int fileNameIndex = txtFile.getName().indexOf(".");
                    String fileName = txtFile.getName().substring(0, fileNameIndex);


                    FileOutputStream fileOut = new FileOutputStream(outExcelFile + fileName + ".xlsx");
                    workbook.write(fileOut);
                    fileOut.close();
                    workbook.close();
                }
            }
        }
    }

    private static List<String[]> readTxtFile(File txtFile) throws IOException {
        List<String[]> data = new ArrayList<>();
        BufferedReader reader = new BufferedReader(new FileReader(txtFile));
        String line;
        int lineCount = 0;
        boolean startLine = false;
        String startString = "Group";

        while ((line = reader.readLine()) != null && lineCount < 13) {
            if (line.contains(startString)) {
                startLine = true;
                String[] rowData = line.split("\\s+");
                data.add(rowData);
                continue;
            }

            if(startLine){
                String[] rowData = line.split("\\s+");
                data.add(rowData);
                lineCount++;
            }
        }

        reader.close();
        return data;
    }
}
