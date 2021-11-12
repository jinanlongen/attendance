import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

public class Attendance {

    public static void main(String[] args) throws Exception {
        Attendance attendance = new Attendance();
//        String model_path = Objects.requireNonNull(attendance.class.getClassLoader().getResource("model.xls")).getFile();
        String model_path = "src/main/resources/model.xls";
        FileInputStream fs = new FileInputStream(model_path);
        POIFSFileSystem ps = new POIFSFileSystem(fs);
        HSSFWorkbook wb = new HSSFWorkbook(ps);
        HSSFSheet sheet = wb.getSheetAt(0);
        HSSFRow row = sheet.getRow(0);
        System.out.println(sheet.getLastRowNum() + "  " + row.getLastCellNum());
        FileOutputStream out = new FileOutputStream("src/main/resources/results.xls");

        HSSFCellStyle color_style = attendance.getStyle(wb, IndexedColors.YELLOW.getIndex());


        row = sheet.getRow((short) (sheet.getLastRowNum() - 1));
        Cell cell1 = row.getCell((short) 2);
        cell1.setCellValue(2000);
//        cell1.setCellStyle(color_style);

        wb.setForceFormulaRecalculation(true);
        out.flush();
        wb.write(out);
        out.close();
        System.out.println(row.getPhysicalNumberOfCells() + "  " + row.getLastCellNum());
    }

    //设置单元格格式
    public HSSFCellStyle getStyle(HSSFWorkbook wb, short color) {
        HSSFCellStyle style = wb.createCellStyle();

        style.setFillForegroundColor(color); // 背景颜色
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        style.setAlignment(HorizontalAlignment.CENTER); // 居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        style.setBorderBottom(BorderStyle.THIN); //下边框
        style.setBorderLeft(BorderStyle.THIN);//左边框
        style.setBorderTop(BorderStyle.THIN);//上边框
        style.setBorderRight(BorderStyle.THIN);//右边框

        return style;
    }

}
