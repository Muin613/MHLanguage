package com.munin.mhlanguage;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

/**
 * Created by Administrator on 2017/11/11.
 */

public class LanguageExcel2Xml {


    static ArrayList<String> name = new ArrayList<>();//指定的name
    static ArrayList<ArrayList<String>> allValues = new ArrayList<>();//各种国家下的name指定的value
    static String srcFileName ="C://Users//Administrator//Desktop//2.xls";//excel的原路径(根据实际修改)
    public static String dstFolderName="C://Users//Administrator//Desktop//values//";// 要生成的文件父路径(根据实际修改)

    public static void main(String[] args) {
        File file = new File(srcFileName);
        try {
            List list = readExcel(file);
            System.out.println(list);
            for (int i = 0; i < list.size(); i++) {
                List one = (List) list.get(i);
                name.add(one.get(0).toString());
                if (i == 0) {
                    for (int innerI = 1; innerI < list.size(); innerI++) {
                        ArrayList<String> value = new ArrayList<>();
                        allValues.add(value);
                    }
                }
                for (int innerI = 0; innerI < list.size() - 1; innerI++) {
                    allValues.get(innerI).add(one.get(innerI + 1).toString());
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }


        for (int i = 0; i < allValues.size(); i++)
            excel2StringXml(name, allValues.get(i));
    }

    public static void excel2StringXml(List<String> name, List<String> values) {
        if (name == null && name.size() <= 1 || values == null && values.size() <= 1)
            return;

        File file = new File(dstFolderName + values.get(0));//文件的所在的文件夹
        if (!file.exists())
            file.mkdirs();

        File txt = new File(dstFolderName + values.get(0) + "//strings.xml");//生成的文件
        if (txt.exists()) {
            txt.delete();
        }


        try {
            txt.createNewFile();
        } catch (IOException e) {
            e.printStackTrace();
        }

        StringBuilder builder = new StringBuilder();
        builder.append("<resources> \n");
        for (int i = 1; i < name.size(); i++)
            builder.append("<string name=\"")
                    .append(name.get(i))
                    .append("\">")
                    .append(values.get(i))
                    .append("</string>\n");
        builder.append("</resources>");

        byte bytes[] = builder.toString().getBytes();
        int b = bytes.length; //

        try {
            FileOutputStream fos = new FileOutputStream(txt);
            fos.write(bytes, 0, b);
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
        }

    }

    public static List<List<Object>> readExcel(File file) throws IOException {
        String fileName = file.getName();
        String extension = fileName.lastIndexOf(".") == -1 ? "" : fileName
                .substring(fileName.lastIndexOf(".") + 1);
        if ("xls".equals(extension)) {
            return read2003Excel(file);
        } else if ("xlsx".equals(extension)) {
            return read2007Excel(file);
        } else {
            throw new IOException("");
        }
    }


    /**
     * Office 2003 excel
     *
     * @throws IOException
     * @throws FileNotFoundException
     */
    private static List<List<Object>> read2003Excel(File file)
            throws IOException {
        List<List<Object>> list = new LinkedList<List<Object>>();
        HSSFWorkbook hwb = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = hwb.getSheetAt(0);
        Object value = null;
        HSSFRow row = null;
        HSSFCell cell = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet
                .getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            List<Object> linked = new LinkedList<Object>();
            for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
                cell = row.getCell(j);
                if (cell == null) {
                    continue;
                }
                DecimalFormat df = new DecimalFormat("0");//
                SimpleDateFormat sdf = new SimpleDateFormat(
                        "yyyy-MM-dd HH:mm:ss");
                DecimalFormat nf = new DecimalFormat("0.00");//
                switch (cell.getCellType()) {
                    case XSSFCell.CELL_TYPE_STRING:
                        value = cell.getStringCellValue();
                        break;
                    case XSSFCell.CELL_TYPE_NUMERIC:
                        if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                            value = df.format(cell.getNumericCellValue());
                        } else if ("General".equals(cell.getCellStyle()
                                .getDataFormatString())) {
                            value = nf.format(cell.getNumericCellValue());
                        } else {
                            value = sdf.format(HSSFDateUtil.getJavaDate(cell
                                    .getNumericCellValue()));
                        }
                        break;
                    case XSSFCell.CELL_TYPE_BOOLEAN:
                        value = cell.getBooleanCellValue();
                        break;
                    case XSSFCell.CELL_TYPE_BLANK:
                        value = "";
                        break;
                    default:
                        value = cell.toString();
                }
                if (value == null || "".equals(value)) {
                    continue;
                }
                linked.add(value);
            }
            list.add(linked);
        }
        return list;
    }

    /**
     * Office 2007 excel
     */
    private static List<List<Object>> read2007Excel(File file)
            throws IOException {
        List<List<Object>> list = new LinkedList<List<Object>>();
        XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet sheet = xwb.getSheetAt(0);
        Object value = null;
        XSSFRow row = null;
        XSSFCell cell = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet
                .getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            List<Object> linked = new LinkedList<Object>();
            for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
                cell = row.getCell(j);
                if (cell == null) {
                    continue;
                }
                DecimalFormat df = new DecimalFormat("0");//
                SimpleDateFormat sdf = new SimpleDateFormat(
                        "yyyy-MM-dd HH:mm:ss");//
                DecimalFormat nf = new DecimalFormat("0.00");//
                switch (cell.getCellType()) {
                    case XSSFCell.CELL_TYPE_STRING:
                        value = cell.getStringCellValue();
                        break;
                    case XSSFCell.CELL_TYPE_NUMERIC:
                        if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                            value = df.format(cell.getNumericCellValue());
                        } else if ("General".equals(cell.getCellStyle()
                                .getDataFormatString())) {
                            value = nf.format(cell.getNumericCellValue());
                        } else {
                            value = sdf.format(HSSFDateUtil.getJavaDate(cell
                                    .getNumericCellValue()));
                        }
                        break;
                    case XSSFCell.CELL_TYPE_BOOLEAN:
                        value = cell.getBooleanCellValue();
                        break;
                    case XSSFCell.CELL_TYPE_BLANK:
                        value = "";
                        break;
                    default:
                        value = cell.toString();
                }
                if (value == null || "".equals(value)) {
                    continue;
                }
                linked.add(value);
            }
            list.add(linked);
        }
        return list;
    }

}
