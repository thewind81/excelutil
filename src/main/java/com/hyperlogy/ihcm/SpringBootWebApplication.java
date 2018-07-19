package com.hyperlogy.ihcm;

import com.hyperlogy.ihcm.excelutil.TestExcelExporter;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.boot.web.servlet.support.SpringBootServletInitializer;
import org.springframework.context.annotation.Bean;
import org.springframework.web.servlet.ViewResolver;
import org.springframework.web.servlet.view.InternalResourceViewResolver;

import java.io.FileNotFoundException;

/**
 * @author Hồ Hữu Trọng
 * @since 17/07/2018
 */
@SpringBootApplication
public class SpringBootWebApplication extends SpringBootServletInitializer {
    @Bean
    public ViewResolver internalResourceViewResolver() {
        InternalResourceViewResolver bean = new InternalResourceViewResolver();
        bean.setPrefix("/WEB-INF/jsp/");
        bean.setSuffix(".jsp");
        return bean;
    }

    @Bean
    public TestExcelExporter testExcelExporter() throws FileNotFoundException {
        return new TestExcelExporter("template.xlsx");
    }

    @Override
    protected SpringApplicationBuilder configure(SpringApplicationBuilder application) {
        return application.sources(SpringBootWebApplication.class);
    }

    public static void main(String[] args) {
        SpringApplication.run(SpringBootWebApplication.class, args);
    }

}