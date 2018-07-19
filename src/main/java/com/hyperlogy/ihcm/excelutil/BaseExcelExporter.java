package com.hyperlogy.ihcm.excelutil;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import java.io.*;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * @param <T>
 */
public abstract class BaseExcelExporter<T> {
    private static final Log logger = LogFactory.getLog(BaseExcelExporter.class);

    protected static final String ENTRY_NAME_SHEET1 = "xl/worksheets/sheet1.xml";
    protected InputStream template;
    protected Writer zipWriter = null;
    protected long startTime;
    protected int currentRow;

    protected BaseExcelExporter(String templatePath) throws FileNotFoundException {
        ClassLoader classloader = Thread.currentThread().getContextClassLoader();
        template = new FileInputStream(new File(classloader.getResource(templatePath).getFile()));
    }

    /**
     * @param outputStream
     * @return
     */
    protected String getXmlSheet1(OutputStream outputStream) throws IOException {
        ZipInputStream zis = new ZipInputStream(template);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ZipOutputStream zos = new ZipOutputStream(outputStream);
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


    public void export(OutputStream outputStream, List<T> objectList) throws IOException {
        startTime = System.currentTimeMillis();
        String xmlSheet1 = getXmlSheet1(outputStream);
        int index = xmlSheet1.indexOf("</sheetData>");
        zipWriter.append(xmlSheet1.substring(0, index));
        // keep hold of the first row (as header), start from second row
        currentRow = 2;
        doExport(objectList);
        zipWriter.append(xmlSheet1.substring(index));
        zipWriter.flush();
        zipWriter.close();
    }

    protected void doExport(List<T> objectList) throws IOException {
        for (int i = 0; i < objectList.size(); i++) {
            writeObject(objectList.get(i));
            zipWriter.flush();
            if (i % 10000 == 0) {
                logger.info("Written " + i + " rows after " + ((System.currentTimeMillis() - startTime) / 1000) + " seconds");
            }
        }
        logger.info("Written " + objectList.size() + " rows after " + ((System.currentTimeMillis() - startTime) / 1000) + " seconds");
    }

    protected abstract void writeObject(T object) throws IOException;

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

    protected void insertRow() throws IOException {
        zipWriter.append("<row r=\"").append(String.valueOf(currentRow)).append("\">");
    }

    protected void endRow() throws IOException {
        zipWriter.append("</row>");
        zipWriter.flush();
        currentRow++;
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
        zipWriter.append("<c r=\"").append(convertNumToColString(columnIndex)).append(String.valueOf(currentRow)).append("\" t=\"inlineStr\" xml:space=\"preserve\"");
        zipWriter.append("><is><t>").append(value).append("</t></is></c>");
    }

    protected void createCell(int columnIndex, int value) throws IOException {
        zipWriter.append("<c r=\"").append(convertNumToColString(columnIndex)).append(String.valueOf(currentRow)).append("\"");
        zipWriter.append("><v>").append(String.valueOf(value)).append("</v></c>");
    }
}
