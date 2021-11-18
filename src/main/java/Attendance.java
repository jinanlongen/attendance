import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.compress.utils.Lists;
import org.apache.commons.io.filefilter.WildcardFileFilter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static javax.xml.bind.JAXBIntrospector.getValue;

public class Attendance {

    public static void main(String[] args) throws Exception {
        String worker = "1543";
        Attendance attendance = new Attendance();
        File dir = new File("src/main/resources");
        FileFilter fileFilter = new WildcardFileFilter("*日报*.xlsx");
        File[] files = dir.listFiles(fileFilter);
        String data_path = "src/main/resources/*日报*.xlsx";
        assert files != null;
        for (File file : files) {
            if (file.getName().startsWith("~"))
            {
                continue;
            }
            data_path = file.getPath();
        }


        XSSFWorkbook data_workbook = attendance.openXSSFFile(data_path);
        XSSFSheet s = data_workbook.getSheetAt(0);
        List<XSSFRow> row_list = Lists.newArrayList();
        for (int rownum=0;rownum<=s.getLastRowNum();rownum++){
            XSSFRow sheetRow = s.getRow(rownum);
            if(sheetRow==null){
                continue;
            }
            String id = sheetRow.getCell(3).getStringCellValue();
            String date = sheetRow.getCell(1).getStringCellValue();
            if (!id.equals("1543")) {
                continue;
            }
            row_list.add(sheetRow);
            //遍历列cell
            for (int cellnum=0;cellnum<=sheetRow.getLastCellNum();cellnum++){
                XSSFCell cell = sheetRow.getCell(cellnum);
                if(cell==null){
                    continue;
                }
                System.out.print( " "+getValue(cell));
            }
            System.out.println();
        }

        String model_path = "src/main/resources/model.xls";
        HSSFWorkbook wb = attendance.openHSSFFile(model_path);

        HSSFSheet sheet = wb.getSheetAt(0);
        HSSFRow row = sheet.getRow(0);
        System.out.println(sheet.getLastRowNum() + "  " + row.getLastCellNum());
        FileOutputStream out = new FileOutputStream("src/main/resources/results.xls");

        HSSFCellStyle color_style = attendance.getStyle(wb, IndexedColors.YELLOW.getIndex());


        row = sheet.getRow((short) (sheet.getLastRowNum() - 1));
        Cell cell1 = row.getCell((short) 2);
        cell1.setCellValue(2000);

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

    public XSSFWorkbook openXSSFFile(String filepath) throws Exception {
        FileInputStream fs = new FileInputStream(filepath);
        if (filepath.endsWith("xlsx")) {
            return new XSSFWorkbook(fs);
    }
        else return null;
    }

    public HSSFWorkbook openHSSFFile(String filepath) throws Exception {
        FileInputStream fs = new FileInputStream(filepath);
        POIFSFileSystem ps = new POIFSFileSystem(fs);
        if (filepath.endsWith("xls")) {
            return new HSSFWorkbook(ps);
        }
        else return null;
    }

}
