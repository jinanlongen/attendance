import java.io.*;
import java.util.List;

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
        String worker = "1599";
        Attendance attendance = new Attendance();
        File dir = new File("src/main/resources");
        FileFilter fileFilter = new WildcardFileFilter("*日报*.xlsx");
        File[] files = dir.listFiles(fileFilter);
        String data_path = "src/main/resources/*日报*.xlsx";
        assert files != null;
        // 查找包含日报的文件位置
        for (File file : files) {
            if (file.getName().startsWith("~")) {
                continue;
            }
            data_path = file.getPath();
        }

        // 打开文件
        XSSFWorkbook data_workbook = attendance.openXSSFFile(data_path);
        XSSFSheet s = data_workbook.getSheetAt(0);
        List<XSSFRow> row_list = Lists.newArrayList();

        // 取出工号为worker的考勤数据
        for (int rownum = 0; rownum <= s.getLastRowNum(); rownum++) {
            XSSFRow sheetRow = s.getRow(rownum);
            if (sheetRow == null) {
                continue;
            }
            String id = sheetRow.getCell(3).getStringCellValue();
            String date = sheetRow.getCell(1).getStringCellValue();
            if (!id.equals(worker)) {
                continue;
            }
            row_list.add(sheetRow);
            //遍历列cell
            for (int cellnum = 0; cellnum <= sheetRow.getLastCellNum(); cellnum++) {
                XSSFCell cell = sheetRow.getCell(cellnum);
                if (cell == null) {
                    continue;
                }
                System.out.print(" " + getValue(cell));
            }
            System.out.println();
        }
        // 打开模板文件
        String model_path = "src/main/resources/model.xls";
        HSSFWorkbook wb = attendance.openHSSFFile(model_path);

        HSSFSheet sheet = wb.getSheetAt(0);
        HSSFRow row = sheet.getRow(0);

        FileOutputStream out = new FileOutputStream("src/main/resources/results.xls");

        // 黄颜色底色表格单元格
        HSSFCellStyle color_style = attendance.getStyle(wb, IndexedColors.YELLOW.getIndex());
        // 无格式单元格
        HSSFCellStyle without_color_style = attendance.getStyle(wb);

        // 考勤记录开始位置
        int row_start_position = 5;
        // 考勤记录结束位置
        int row_end_position = 5 + row_list.size();

        for (int i = row_list.size() - 1; i >= 0; i--) {
            XSSFRow row_data = row_list.get(i);
            String day = row_data.getCell(1).getStringCellValue();
            // 设置格式
            if (day.equals("星期六") || day.equals("星期日")) {
                sheet.getRow(row_start_position).getCell(2).setCellStyle(color_style);
                sheet.getRow(row_start_position).getCell(3).setCellStyle(color_style);
                sheet.getRow(row_start_position).getCell(4).setCellStyle(color_style);
                sheet.getRow(row_start_position).getCell(5).setCellStyle(color_style);
                sheet.getRow(row_start_position).getCell(6).setCellStyle(color_style);
            }
            sheet.getRow(row_start_position).getCell(7).setCellStyle(without_color_style);
            sheet.getRow(row_start_position).getCell(8).setCellStyle(without_color_style);
            sheet.getRow(row_start_position).getCell(9).setCellStyle(without_color_style);
            sheet.getRow(row_start_position).getCell(10).setCellStyle(without_color_style);
            sheet.getRow(row_start_position).getCell(11).setCellStyle(without_color_style);
            sheet.getRow(row_start_position).getCell(12).setCellStyle(without_color_style);
            sheet.getRow(row_start_position).getCell(13).setCellStyle(without_color_style);
            sheet.getRow(row_start_position).getCell(14).setCellStyle(without_color_style);
            sheet.getRow(row_start_position).getCell(15).setCellStyle(without_color_style);

            //日期
            sheet.getRow(row_start_position).getCell(7).setCellValue(row_data.getCell(0).getStringCellValue());

            // 星期
            sheet.getRow(row_start_position).getCell(8).setCellValue(row_data.getCell(1).getStringCellValue());

            // 姓名
            sheet.getRow(row_start_position).getCell(9).setCellValue(row_data.getCell(2).getStringCellValue());

            //最早
            sheet.getRow(row_start_position).getCell(10).setCellValue(row_data.getCell(7).getStringCellValue());

            //最晚
            sheet.getRow(row_start_position).getCell(11).setCellValue(row_data.getCell(8).getStringCellValue());
            //计算加班时长
            String latest = row_data.getCell(5).getStringCellValue();
//            if (!latest.equals("--")){
//                String[] nns = latest.split(":");
//                int hour = Integer.parseInt(nns[0]);
//                int min = Integer.parseInt(nns[1]);
//                if (hour - 18 > 1)
//                {
//                    if (min -30 >=0)
//                    {
//                        double minutes = 0.5;
//                    }
//                }
//
//            }

            // 打卡次数
            try {
                sheet.getRow(row_start_position).getCell(12).setCellValue(row_data.getCell(9).getNumericCellValue());
            }
            catch (Exception e){
                sheet.getRow(row_start_position).getCell(12).setCellValue(row_data.getCell(9).getStringCellValue());
            }

            // 时长
            sheet.getRow(row_start_position).getCell(13).setCellValue(row_data.getCell(11).getStringCellValue());

            // 详细
            sheet.getRow(row_start_position).getCell(14).setCellValue(row_data.getCell(13).getStringCellValue());

            // 假勤申请
            sheet.getRow(row_start_position).getCell(15).setCellValue(row_data.getCell(12).getStringCellValue());

            row_start_position +=1;
        }

//        row = sheet.getRow(row_start_position);
//        Cell cell1 = row.getCell((short) 2);
//        cell1.setCellValue(2000);

        wb.setForceFormulaRecalculation(true);
        out.flush();
        wb.write(out);
        out.close();
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

    //设置单元格格式
    public HSSFCellStyle getStyle(HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();

        style.setAlignment(HorizontalAlignment.CENTER); // 居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);
//
//        style.setBorderBottom(BorderStyle.THIN); //下边框
//        style.setBorderLeft(BorderStyle.THIN);//左边框
//        style.setBorderTop(BorderStyle.THIN);//上边框
//        style.setBorderRight(BorderStyle.THIN);//右边框

        return style;
    }

    public XSSFWorkbook openXSSFFile(String filepath) throws Exception {
        FileInputStream fs = new FileInputStream(filepath);
        if (filepath.endsWith("xlsx")) {
            return new XSSFWorkbook(fs);
        } else return null;
    }

    public HSSFWorkbook openHSSFFile(String filepath) throws Exception {
        FileInputStream fs = new FileInputStream(filepath);
        POIFSFileSystem ps = new POIFSFileSystem(fs);
        if (filepath.endsWith("xls")) {
            return new HSSFWorkbook(ps);
        } else return null;
    }

}
