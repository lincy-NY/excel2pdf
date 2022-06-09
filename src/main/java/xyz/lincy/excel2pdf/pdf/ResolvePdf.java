package xyz.lincy.excel2pdf.pdf;

import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;

public class ResolvePdf {

    public void toPdf(String filePath){
        //加载Excel文档
        Workbook wb = new Workbook();
        wb.loadFromFile(filePath);

        //调用方法保存为PDF格式
        wb.saveToFile("/Users/lincy/ToPDF.pdf", FileFormat.PDF);
//        wb.saveToFile("ToPDF.pdf",FileFormat.PDF);
    }
}
