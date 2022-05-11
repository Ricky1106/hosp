package btg;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;

public class BtgMain {

    public static void main(String[] args) {
        System.out.println("btg main start");

        ReadBtgFile readBtgFile = new ReadBtgFile();
//        String path = "D:\\Downloads\\Transactionscsv.csv";
        String path = "D:\\Downloads\\Transactionsxlsx.xlsx";
        Workbook workbook = null;
        try {
            workbook = readBtgFile.getWorkbook(path);
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("workbook:" + workbook);


        OutputBtgExcelFile outputBtgExcelFile = new OutputBtgExcelFile();
        try {
            outputBtgExcelFile.outputExcelFile(workbook,"D:\\Downloads\\excelOutput");
        } catch (IOException e) {
            e.printStackTrace();
        }

    }


}
