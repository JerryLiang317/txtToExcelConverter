package org.example;

import java.io.IOException;
import java.util.Scanner;

import static org.example.TxtToExcelMethod.txtToExcel;

public class Main {
    public static void main(String[] args) throws IOException {
        Scanner input = new Scanner(System.in);
        System.out.print("請輸入想要讀取的檔案位置 (路徑): ");
        String filePath = input.nextLine();
        System.out.print("請輸入要產生 Excel 檔的路徑位置： ");
        String outExcelFile = input.nextLine();

        txtToExcel(filePath, outExcelFile);

        System.out.println("Excel 寫入完成");
    }
}