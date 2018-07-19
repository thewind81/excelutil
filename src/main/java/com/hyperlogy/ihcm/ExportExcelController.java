package com.hyperlogy.ihcm;

import com.hyperlogy.ihcm.excelutil.TestExcelExporter;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

/**
 * @author Hồ Hữu Trọng
 * @since 17/07/2018
 */
@Controller
public class ExportExcelController {

    @Autowired
    private TestExcelExporter testExcelExporter;

    @RequestMapping("/")
    public String welcome() {
        return "welcome";
    }

    @RequestMapping("/export")
    public void export(HttpServletResponse response, @RequestParam int rowCount) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=\"Happy Nice Day.xlsx\"");
        testExcelExporter.setRowCount(rowCount);
        testExcelExporter.export(response.getOutputStream(), null);
    }
}
