package com.example.excel.service;

import com.example.excel.util.ExportExcel;
import com.example.excel.util.ImportExcel;
import com.example.excel.util.JsonUtil;
import com.fasterxml.jackson.databind.JsonNode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

/**
 * @author ClowLAY
 * create date 2020/3/25
 */
@Service
public class ExcelService {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelService.class);

    //excel标题
    private String title = "";

    //表格数据
    private List<Object[]> dataList;

    //excel列头信息
    private String[] colsName;
    /**
     * 导出excel
     */
    public void exportExcel(HttpServletResponse response){

        title = "导入标题";

        dataList = new ArrayList<Object[]>();

        colsName = new String[] { "书籍编号", "借书人", "借书时间", "预计还书时间", "借书状态" };

        for (int i = 0; i < 6; i++) {
            var objects = new Object[colsName.length];
            objects[0] = "HK200"+i;
            objects[1] = "小明"+i;
            Timestamp currentTimestamp=Timestamp.valueOf(LocalDateTime.now());
            objects[2] = currentTimestamp;
            objects[3] = "未定";
            if (i%2 == 0) {
                objects[4] = "借书中";
            }else {
                objects[4] = "已还书";
            }

            dataList.add(objects);

        }

        try {
            response.reset();
            response.setContentType("application/octet-stream; charset=utf-8");
            response.setHeader("Content-disposition", "attachment;filename="+ URLEncoder.encode("导出表名.xlsx", "UTF-8"));
            var out = response.getOutputStream();
            var exportExcel = new ExportExcel(title, colsName, dataList);
            exportExcel.export(out);

            out.flush();
            out.close();
            LOGGER.info("导出Excel成功");
        } catch (IOException e) {
            e.printStackTrace();
            LOGGER.info("导出Excel失败");
        }

    }

    public JsonNode importExcel(MultipartFile file){
        var root= JsonUtil.OBJECT_MAPPER.createObjectNode();
        root.put("result","fail");
        if (file.isEmpty()){
            root.put("cause","必要参数未找到");
            return root;
        }
        ImportExcel excel;
        try {
            excel=new ImportExcel(file);
        }catch (Exception e){
            root.put("cause","文件识别失败");
            return root;
        }

        try {
            dataList=excel.getColNames();
        } catch (IOException e) {
            e.printStackTrace();
            root.put("cause","数据解析失败");
            return root;
        }
        dataList.forEach(data->{
            for (var i=0;i<data.length-1;i++){
                System.out.print(data[i]+"\t");
            }

            System.out.println();
        });
        root.put("result","success");
        return root;
    }


}

