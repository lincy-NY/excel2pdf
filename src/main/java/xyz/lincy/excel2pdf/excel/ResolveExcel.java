package xyz.lincy.excel2pdf.excel;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

// 解决excel相关的内容，读取excel、调整行高
public class ResolveExcel {

    public void readExcel() throws Exception{
        // 创建 Excel 文件的输入流对象
        FileInputStream excelFileInputStream = new FileInputStream(ResolveExcel.class.getClassLoader().getResource("1.xls").getFile());
//        // XSSFWorkbook 就代表一个 Excel 文件
//        // 创建其对象，就打开这个 Excel 文件
//        XSSFWorkbook workbook = new XSSFWorkbook(excelFileInputStream);
//        // 输入流使用后，及时关闭！这是文件流操作中极好的一个习惯！
//        excelFileInputStream.close();
//        // XSSFSheet 代表 Excel 文件中的一张表格
//        // 我们通过 getSheetAt(0) 指定表格索引来获取对应表格
//        // 注意表格索引从 0 开始！
//        XSSFSheet sheet = workbook.getSheetAt(0);

//        HSSFWorkbook wb = new HSSFWorkbook(excelFileInputStream);
//        HSSFSheet sheet = wb.getSheetAt(0);

        Workbook rwb = WorkbookFactory.create(excelFileInputStream);
        Sheet sheet1 = rwb.getSheetAt(0);
        Row row = sheet1.getRow(3);
        Cell cell = row.getCell(0);
        System.out.println(cell.getStringCellValue());

    }
}
