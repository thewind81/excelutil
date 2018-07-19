package com.hyperlogy.ihcm.excelutil;

import java.io.FileNotFoundException;
import java.io.IOException;

public class SimpleExcelExporter extends BaseExcelExporter {
    public SimpleExcelExporter(String templatePath) throws FileNotFoundException {
        super(templatePath);
    }

    @Override
    protected void writeObject(Object object) throws IOException {
    }
}
