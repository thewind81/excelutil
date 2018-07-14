package com.hyperlogy.ihcm.excelutil;

import java.io.*;
import java.util.Random;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class SimpleExcelExporter {
    protected static final String ENTRY_NAME_SHEET1 = "xl/worksheets/sheet1.xml";
    protected static final String NAME = "abcdefghijklmnopqrstuvwxyz";
    protected static final Random RANDOM = new Random();
    protected InputStream template;
    protected Writer zipWriter = null;
    private int currentRow;

    public SimpleExcelExporter(String templatePath) throws FileNotFoundException {
        ClassLoader classloader = Thread.currentThread().getContextClassLoader();
        template = new FileInputStream(new File(classloader.getResource(templatePath).getFile()));
    }

    protected String getXmlSheet1(InputStream outputTemplate, OutputStream os) throws IOException {
        ZipInputStream zis = new ZipInputStream(outputTemplate);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ZipOutputStream zos = new ZipOutputStream(os);
        ZipEntry ze = null;
        while ((ze = zis.getNextEntry()) != null) {
            if (ze.getName().equals(ENTRY_NAME_SHEET1)) {
                copyStream(zis, baos);
            } else {
                copyZip(ze.getName(), zis, zos);
            }
        }
        zis.close();
        zos.putNextEntry(new ZipEntry(ENTRY_NAME_SHEET1));
        zipWriter = new OutputStreamWriter(zos, "UTF-8");
        return baos.toString();
    }

    public void export(OutputStream os, int rowCount) throws IOException {
        long startTime = System.currentTimeMillis();
        String xmlSheet1 = getXmlSheet1(template, os);
        int index = xmlSheet1.indexOf("</sheetData>");
        zipWriter.append(xmlSheet1.substring(0, index));
        int rowIndex = 1;
        for (int i = 0; i < rowCount; i++) {
            writeObject(rowIndex);
            rowIndex++;
            zipWriter.flush();
            if (i % 10000 == 0) {
                System.out.println("Written " + i + " rows after " + ((System.currentTimeMillis() - startTime) / 1000) + " seconds");
            }
        }
        zipWriter.append(xmlSheet1.substring(index));
        zipWriter.flush();
        zipWriter.close();
        System.out.println("Written " + rowCount + " rows after " + ((System.currentTimeMillis() - startTime) / 1000) + " seconds");
    }


    protected void writeObject(int rowIndex) throws IOException {
        insertRow(rowIndex);
        createCell(0, rowIndex);
        createCell(1, randomName());
        createCell(2, randomPhone());
        createCell(3, randomEmail());
        endRow();
    }

    private String randomName() {
        return randomString(true) + " " + randomString(true);
    }

    private String randomString(boolean upperFirst) {
        StringBuilder sb = new StringBuilder();
        int length = RANDOM.nextInt(5) + 3;
        for (int i = 0; i < length; i++) {
            if (i == 0 && upperFirst) {
                sb.append(Character.toUpperCase(NAME.charAt(RANDOM.nextInt(NAME.length()))));
            } else {
                sb.append(NAME.charAt(RANDOM.nextInt(NAME.length())));
            }
        }
        return sb.toString();
    }

    private String randomPhone() {
        StringBuilder sb = new StringBuilder();
        sb.append("0");
        for (int i = 0; i < 8; i++) {
            sb.append(RANDOM.nextInt(9));
        }
        return sb.toString();
    }

    private String randomEmail() {
        return randomString(false) + "@hyperlogy.com";
    }

    protected void copyZip(String entryName, ZipInputStream zis, ZipOutputStream zos) throws IOException {
        zos.putNextEntry(new ZipEntry(entryName));
        copyStream(zis, zos);
        zos.flush();
    }

    protected void copyStream(InputStream in, OutputStream out) throws IOException {
        byte[] chunk = new byte[1024];
        int count;
        while ((count = in.read(chunk)) >= 0) {
            out.write(chunk, 0, count);
        }
    }

    protected void insertRow(int rownum) throws IOException {
        zipWriter.append("<row r=\"").append(String.valueOf(rownum + 1)).append("\">");
        this.currentRow = rownum;
    }

    protected void endRow() throws IOException {
        zipWriter.append("</row>");
        zipWriter.flush();
    }

    protected String convertNumToColString(int col) {
        int excelColNum = col + 1;
        String colRef = "";
        int colRemain = excelColNum;
        while (colRemain > 0) {
            int thisPart = colRemain % 26;
            if (thisPart == 0) {
                thisPart = 26;
            }
            colRemain = (colRemain - thisPart) / 26;
            char colChar = (char) (thisPart + 64);
            colRef = colChar + colRef;
        }
        return colRef;
    }

    protected void createCell(int columnIndex, String value) throws IOException {
        zipWriter.append("<c r=\"").append(convertNumToColString(columnIndex)).append(String.valueOf(currentRow + 1)).append("\" t=\"inlineStr\" xml:space=\"preserve\"");
        zipWriter.append("><is><t>").append(value).append("</t></is></c>");
    }

    protected void createCell(int columnIndex, int value) throws IOException {
        zipWriter.append("<c r=\"").append(convertNumToColString(columnIndex)).append(String.valueOf(currentRow + 1)).append("\"");
        zipWriter.append("><v>").append(String.valueOf(value)).append("</v></c>");
    }
}
