package xyz.lincy.excel2pdf;

import org.apache.poi.ss.usermodel.Workbook;
import xyz.lincy.excel2pdf.excel.ResolveExcel;
import xyz.lincy.excel2pdf.pdf.ResolvePdf;

import java.io.FileOutputStream;

public class Main {

    public static void main(String[] args) throws Exception {
        ResolveExcel re = new ResolveExcel();
        Workbook wb = re.readExcel("/Users/lincy/2.xlsx");
        re.changeRowHeight(wb, 4, 40);
        re.changeRowHeight(wb, 14, 80);
        re.changeRowHeight(wb, 25, 40);
        wb.write(new FileOutputStream("/Users/lincy/3.xlsx"));
        ResolvePdf rp = new ResolvePdf();
        rp.toPdf("/Users/lincy/3.xlsx");
    }
}
