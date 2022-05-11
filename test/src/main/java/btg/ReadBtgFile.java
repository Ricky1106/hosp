package btg;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;


public class ReadBtgFile {


    public Workbook getWorkbook(String path) throws FileNotFoundException, IOException {

        System.out.println("read btg file start ");

        Workbook wb = null;
        if (path == null) return null;
        String extString = path.substring(path.lastIndexOf(".")).replace(".","");
        System.out.println("extString:" + extString);
        InputStream is = new FileInputStream(path);
        if ("XLS".equalsIgnoreCase(extString)) {
            System.out.println("file type is xls");
            wb = new HSSFWorkbook(is);
        } else if ("XLSX".equalsIgnoreCase(extString)) {
            System.out.println("file type is xlsx");
            wb = new XSSFWorkbook(is);
        } else if ("CSV".equalsIgnoreCase(extString)) {
            System.out.println("file type is csv");
            readCsv(path);
        }
        return wb;
    }


    private void readCsv(String path) {
        try {
            InputStreamReader isr = new InputStreamReader(new FileInputStream(path));//檔案讀取路徑
            BufferedReader reader = new BufferedReader(isr);
//            BufferedWriter bw = new BufferedWriter(new FileWriter(path));//檔案輸出路徑
            String line = null;
            while ((line = reader.readLine()) != null) {
                String item[] = line.split(",");
                /** 讀取 **/
                String data1 = item[0].trim();
                String data2 = item[1].trim();
                String data3 = item[2].trim();

                System.out.print(data1 + "\t" + data2 + "\t" + data3 + "\n");

                //可自行變化成存入陣列或arrayList方便之後存取
            }
//            bw.close();
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


}
