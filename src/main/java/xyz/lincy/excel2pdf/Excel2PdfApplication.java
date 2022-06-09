package xyz.lincy.excel2pdf;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import xyz.lincy.excel2pdf.excel.ResolveExcel;

import java.io.FileOutputStream;

@SpringBootApplication
public class Excel2PdfApplication implements ApplicationRunner{

    public static void main(String[] args) {
        SpringApplication.run(Excel2PdfApplication.class, args);
    }

    @Override
    public void run(ApplicationArguments args) throws Exception {
        ResolveExcel re = new ResolveExcel();
        Workbook wb = re.readExcel("/Users/lincy/2.xlsx");
        wb.write(new FileOutputStream("/Users/lincy/3.xlsx"));
    }
}
