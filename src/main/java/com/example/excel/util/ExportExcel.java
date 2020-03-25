package com.example.excel.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * excel导出工具类
 *
 * @author ClowLAY
 * create date 2020/3/25
 */
public class ExportExcel {
    //导出表的标题
    private String title;

    //导出表的列名
    private String[] colName;

    private List<Object[]> dataList=new ArrayList<>();

    //构造函数。传入要导入的数据
    public ExportExcel(String title, String[] colName, List<Object[]> dataList) {
        this.title = title;
        this.colName = colName;
        this.dataList = dataList;
    }


    /**
     * 导出数据
     */
    public void export(OutputStream out){

        //工作薄对象
        Workbook workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet(title);

        //产生表格标题行
        var rowTitle = sheet.createRow(0);
        rowTitle.setHeightInPoints(30);
        var cellTitle = rowTitle.createCell(0);

        //表格样式地定义
        var getTitleTopStyle = this.getTitleTopStyle(workbook);
        var columnTopStyle = this.getColumnTopStyle(workbook);
        var style = this.getStyle(workbook);

        cellTitle.setCellStyle(getTitleTopStyle);
        cellTitle.setCellValue(title);
        sheet.addMergedRegion(new CellRangeAddress(0,
                0, 0, colName.length-1));
        //定义所需列数
        var columnNum = colName.length;
        var rowRowNmae = sheet.createRow(1);

        //将列头设置到表格的单元格中
        for (var i = 0; i < columnNum; i++) {
            var cellRowName = rowRowNmae.createCell(i);
            cellRowName.setCellType(Cell.CELL_TYPE_STRING);
            RichTextString text = new XSSFRichTextString(colName[i]);
            cellRowName.setCellValue(text);
            cellRowName.setCellStyle(columnTopStyle);

        }

        //将查询的数据设置到sheet对应的单元格中
        if (dataList.size()>0){
            for (var i = 0; i < dataList.size(); i++) {
                var objects = dataList.get(i);//遍历每个对象
                var row = sheet.createRow(i + 2);

                for (var a = 0; a < objects.length; a++) {

                    var cell = row.createCell(a, HSSFCell.CELL_TYPE_STRING);
                    if (!"".equals(objects[a]) && objects[a] != null) {
                        cell.setCellValue(objects[a].toString());
                    } else {
                        cell.setCellValue("");
                    }
                    cell.setCellStyle(style);

                }
            }
        }


        //让列宽随着导出的列长自动适应
        for (var colNum = 0; colNum < columnNum; colNum++) {
            var columnWidth = sheet.getColumnWidth(colNum) / 256;
            for (var rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++){
                Row currentRow;
                if (sheet.getRow(rowNum) == null){
                    currentRow=sheet.createRow(rowNum);
                } else {
                    currentRow = sheet.getRow(rowNum);
                }

                if (currentRow.getCell(colNum) != null) {
                    Cell currentCell = currentRow.getCell(colNum);
                    if (currentCell.getCellType() ==Cell.CELL_TYPE_STRING) {
                        var length=0;
                        if (currentCell.getStringCellValue()!=null){
                            length = currentCell.getStringCellValue().getBytes().length;
                        }

                        if (columnWidth < length) {
                            columnWidth = length;
                        }
                    }
                }
            }

            sheet.setColumnWidth(colNum,(columnWidth + 4) * 256);

        }
        if (workbook != null){
            try {
                workbook.write(out);

            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /*
     *列头单元格样式
     */
    public CellStyle getColumnTopStyle(Workbook workbook) {

        var styles = new HashMap<String, CellStyle>();
        // 设置样式
        CellStyle style = workbook.createCellStyle();
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font dataFont = workbook.createFont();
        dataFont.setFontName("Arial");
        dataFont.setFontHeightInPoints((short) 10);
        style.setFont(dataFont);
        styles.put("data", style);

        // 设置样式
        style = workbook.createCellStyle();


        style.cloneStyleFrom(styles.get("data"));
//		style.setWrapText(true);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        // 设置字体
        Font headerFont = workbook.createFont();
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(headerFont);
        return style;

    }

    /*
     *标题单元格样式
     */
    public CellStyle getTitleTopStyle(Workbook workbook) {
        // 设置字体
        Font font = workbook.createFont();

        // 设置字体大小
        font.setFontHeightInPoints((short) 13);
        // 字体加粗
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        // 设置字体名字
        font.setFontName("Arial");
        // 设置样式
        CellStyle style = workbook.createCellStyle();
        //设置单元格的水平对齐类型
        style.setAlignment(CellStyle.ALIGN_CENTER);

        //设置单元格的垂直对齐类型
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        // 设置低边框
        style.setBorderBottom(CellStyle.BORDER_THIN);
        // 设置低边框颜色
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        // 设置右边框
        style.setBorderRight(CellStyle.BORDER_THIN);
        // 设置顶边框
        //style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        // 设置顶边框颜色
        //style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        // 在样式中应用设置的字体
        style.setFont(font);
        // 设置自动换行
        style.setWrapText(false);
        // 设置水平对齐的样式为居中对齐；
        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        return style;

    }

    public CellStyle getStyle(Workbook workbook) {
        // 设置字体
        Font font = workbook.createFont();
        // 设置字体大小
        font.setFontHeightInPoints((short) 10);
        // 字体加粗
        //font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        // 设置字体名字
        font.setFontName("Arial");
        // 设置样式;
        CellStyle style = workbook.createCellStyle();
        // 设置底边框;
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        // 设置底边框颜色;
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        // 设置左边框;
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        // 设置左边框颜色;
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        // 设置右边框;
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);
        // 设置右边框颜色;
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        // 设置顶边框;
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        // 设置顶边框颜色;
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        // 在样式用应用设置的字体;
        style.setFont(font);
        // 设置自动换行;
        style.setWrapText(false);
        // 设置水平对齐的样式为居中对齐;
        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        // 设置垂直对齐的样式为居中对齐;
        style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        return style;
    }



}
