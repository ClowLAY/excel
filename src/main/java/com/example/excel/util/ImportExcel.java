package com.example.excel.util;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * 导入excel内容的工具类
 * @author ClowLAY
 * create date 2020/3/25
 */
public class ImportExcel {

    //文件对象
    private MultipartFile file;

    /**
     * 工作薄对象
     */
    private Workbook workbook;

    public ImportExcel(MultipartFile file){
        this.file = file;
    }

    /**
     * 返回除标题的外的所有数据，第一个对象数组为列数据
     * @return
     */
    public List<Object[]> getColNames() throws IOException {
        var dataList=new ArrayList<Object[]>();

        var rowSize=0;

        //根据excel表格式新建表对象
        var in = new BufferedInputStream(file.getInputStream());
        if (StringUtils.isBlank(file.getOriginalFilename())) {
            throw new IOException("Import file is empty!");
        } else if (file.getOriginalFilename().toLowerCase().endsWith("xls")) {
            this.workbook = new HSSFWorkbook(in);
        } else if (file.getOriginalFilename().toLowerCase().endsWith("xlsx")) {
            this.workbook = new XSSFWorkbook(in);
        } else {
            throw new IOException("Invalid import file type!");
        }

        Cell cell=null;
        for (var sheetIndex=0;sheetIndex<workbook.getNumberOfSheets();sheetIndex++){
            //获取指定索引的表对象
            var sheetAt = workbook.getSheetAt(sheetIndex);

            //第一行作为标题，不取
            for (var rowIndex=1;rowIndex<=sheetAt.getLastRowNum();rowIndex++) {
                var row=sheetAt.getRow(rowIndex);
                if (row == null){
                    continue;
                }
                //获取一列单元格数量
                var tempRowSize  = row.getLastCellNum();
                if (tempRowSize > rowSize){
                    rowSize = tempRowSize;
                }

                var values=new Object[rowSize];
                boolean hasValue = false;
                for (var columnIndex = 0 ;columnIndex < row.getLastCellNum(); columnIndex++ ){
                    Object value="";
                    cell = row.getCell(columnIndex);
                    if (cell != null){
                        switch (cell.getCellType()){
                            case Cell.CELL_TYPE_STRING :
                                value = cell.getStringCellValue();
                                break;
                            case Cell.CELL_TYPE_NUMERIC :
                                if (HSSFDateUtil.isCellDateFormatted(cell)){
                                    var date=cell.getDateCellValue();
                                    if (date != null){
                                        value = new SimpleDateFormat("yyyy-MM-dd").format(date);
                                    }else
                                        value="";
                                } else {
                                    value= new DecimalFormat("0").format(cell.getNumericCellValue());
                                }
                                break;
                            case Cell.CELL_TYPE_FORMULA :
                                value=cell.getCellFormula();
                                break;
                            case Cell.CELL_TYPE_BOOLEAN :
                                value=cell.getBooleanCellValue();
                                break;
                            case Cell.CELL_TYPE_BLANK :
                                value="";
                                break;
                            case Cell.CELL_TYPE_ERROR :
                                value=cell.getErrorCellValue();
                                break;
                            default :
                                value="";
                        }
                    }
                    //如果当前行的第一列为空，则跳过
                    if (columnIndex == 0 && value.toString().trim().equals("")) {
                        break;
                    }
                    values[columnIndex] = value.toString().trim();
                    hasValue = true;
                }
                if (hasValue) {
                    dataList.add(values);
                }

            }

        }
        in.close();

        return dataList;
    }

    /**
     * 导入测试
     */
	/*public static void main(String[] args) throws Throwable {

		ImportExcel excel = new ImportExcel(new File("D:/2020-03-23.xlsx"));

		var dataList=excel.getColNames();

		System.out.println("size="+dataList.size());
		dataList.forEach(data->{
		    for (var i=0;i<data.length-1;i++){
                System.out.print(i+","+data[i]+"\t");
            }

            System.out.println();
        });

	}*/

}
