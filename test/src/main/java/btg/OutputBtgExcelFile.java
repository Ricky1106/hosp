package btg;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Locale;

public class OutputBtgExcelFile {

    /**
     * @param path
     * @throws IOException
     */
    public void outputExcelFile(Workbook workbook, String path) throws IOException {
        FileOutputStream fOut = null;
        try {
            LocalDateTime now = LocalDateTime.now();
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
            String formatDateTime = now.format(formatter);
            String fileName = "btg_"+formatDateTime;

            System.out.println("output:"+path+"\\"+fileName);
            fOut = new FileOutputStream(path+"\\"+fileName);

            workbook.write(fOut);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        // 把相應的Excel 工作簿存檔
        catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } finally {
            if (fOut != null) {
                fOut.flush();
                // 操作結束，關閉檔案
                System.out.println("檔案生成...");
                fOut.close();
            }

        }
    }
}
