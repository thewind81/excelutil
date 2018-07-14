package com.hyperlogy.ihcm;

import com.hyperlogy.ihcm.excelutil.SimpleExcelExporter;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@Controller
public class ExportExcelController {

    @Autowired
    private SimpleExcelExporter excelExporter;

    @RequestMapping("/")
    public String welcome() {
        return "welcome";
    }

    @RequestMapping("/export")
    public void export(HttpServletResponse response, @RequestParam int rowCount) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=\"Happy Nice Day.xlsx\"");
        excelExporter.export(response.getOutputStream(), rowCount);
    }
}
