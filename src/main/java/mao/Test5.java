package mao;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

/**
 * Project name(项目名称)：java报表_POI格式设置
 * Package(包名): mao
 * Class(类名): Test5
 * Author(作者）: mao
 * Author QQ：1296193245
 * GitHub：https://github.com/maomao124/
 * Date(创建日期)： 2023/6/3
 * Time(创建时间)： 13:06
 * Version(版本): 1.0
 * Description(描述)： 无
 */

public class Test5
{
    public static void main(String[] args)
    {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("test");
        Font font = workbook.createFont();
        //加粗
        font.setBold(true);
        //字体名称
        font.setFontName("黑体");
        //字体颜色
        font.setColor(Font.COLOR_RED);
        //字体大小
        font.setFontHeightInPoints((short) 20);
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("test1");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(1);
        cell.setCellValue("test2");
        cell.setCellStyle(cellStyle);

        font = workbook.createFont();
        //加粗
        font.setBold(true);
        //字体名称
        font.setFontName("宋体");
        //字体颜色
        font.setColor((short) 11);
        //字体大小
        font.setFontHeightInPoints((short) 25);
        cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);

        cell = row.createCell(2);
        cell.setCellValue("test3");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(3);
        cell.setCellValue("test4");
        cell.setCellStyle(cellStyle);

        font = workbook.createFont();
        //加粗
        font.setBold(false);
        //字体名称
        font.setFontName("微软雅黑");
        //字体颜色
        font.setColor((short) 12);
        //字体大小
        font.setFontHeightInPoints((short) 12);
        cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);

        cell = row.createCell(4);
        cell.setCellValue("test5");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(5);
        cell.setCellValue("test6");
        cell.setCellStyle(cellStyle);

        try (FileOutputStream fileOutputStream = new FileOutputStream("./out5.xlsx"))
        {
            workbook.write(fileOutputStream);
            workbook.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}
