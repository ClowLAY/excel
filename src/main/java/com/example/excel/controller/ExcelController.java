package com.example.excel.controller;

import com.example.excel.service.ExcelService;
import com.fasterxml.jackson.databind.JsonNode;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;

/**
 * @author ClowLAY
 * create date 2020/3/25
 */
@Controller
public class ExcelController {

    private ExcelService excelService;;

    @Autowired
    public ExcelController(ExcelService excelService) {
        this.excelService = excelService;
    }

    @GetMapping(value = "export",produces = "application/json;charset=UTF-8")
    public void export(HttpServletResponse response)  {
        excelService.exportExcel(response);
    }

    @PostMapping(value = "import",produces = "application/json;charset=UTF-8")
    @ResponseBody
    public JsonNode importAsset(@RequestParam("file") MultipartFile multipartFile) {
        return  excelService.importExcel(multipartFile);
    }


    @GetMapping("/index")
    public String index() {
        return "web/index";
    }


}
