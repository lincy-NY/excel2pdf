package xyz.lincy.excel2pdf.excel;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;

// 解决excel相关的内容，读取excel、调整行高
public class ResolveExcel {

    /**
     * 读取excel
     * @return
     * @throws Exception
     */
    public Workbook readExcel(String filePath) throws Exception{
        // 创建 Excel 文件的输入流对象
        FileInputStream excelFileInputStream = new FileInputStream(filePath);

        Workbook wb = WorkbookFactory.create(excelFileInputStream);
//        Workbook rwb = WorkbookFactory.create(excelFileInputStream);
        return wb;
    }

    /**
     * 改变某一行的行高
     * @param rowNum
     * @return
     */
    public Workbook changeRowHeight(Workbook wb, int rowNum, int height){
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(rowNum);
        // 获取第0列
        Cell cell = row.getCell(0);
        int length = cell.getStringCellValue().length();
        row.setHeightInPoints(height);
        return wb;
    }

    //将excel写回文件
    public void writeExcel(Workbook wb, String filePath) throws Exception{
        FileOutputStream outputStream = new FileOutputStream(filePath);
        wb.write(outputStream);
    }


}
