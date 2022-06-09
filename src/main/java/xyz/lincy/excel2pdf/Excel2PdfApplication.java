package xyz.lincy.excel2pdf;

import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import xyz.lincy.excel2pdf.excel.ResolveExcel;

@SpringBootApplication
public class Excel2PdfApplication implements ApplicationRunner{

    public static void main(String[] args) {
        SpringApplication.run(Excel2PdfApplication.class, args);
    }

    @Override
    public void run(ApplicationArguments args) throws Exception {
        new ResolveExcel().readExcel();
    }
}
