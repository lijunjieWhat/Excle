package com.yuanjun;

import com.alibaba.excel.EasyExcel;

import com.alibaba.fastjson.JSON;
import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.sound.midi.Soundbank;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Scanner;

import static org.apache.poi.ss.usermodel.CellType.*;

/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) throws Exception {

        System.out.println("请将送货单放入-D:\\excelOption 目录下");
        System.out.println("对账单模板放入-D:\\模板\\对账单 目录下");
        Scanner scanner = new Scanner(System.in);
        System.out.println("是否准备好？输入1开始运行");
        int flag = scanner.nextInt();
        if (flag == 1) {
            List<DemoData> list = getExcelValue();
            String fileName = "D:\\模板\\对账单\\2022.xlsx";
            FileInputStream excelFileInputStream = new FileInputStream(fileName);
            XSSFWorkbook workbook = new XSSFWorkbook(excelFileInputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);

            sheet.shiftRows(8, sheet.getLastRowNum(), list.size());

            for (int i = 0; i < list.size(); i++) {
                sheet.copyRows(7, 7, 8 + i, new CellCopyPolicy());


            }
            int sum = 0;
            for (int i = 0; i < list.size(); i++) {
                sheet.getRow(7 + i).getCell(0).setCellValue(list.get(i).getId());
                sheet.getRow(7 + i).getCell(1).setCellValue(list.get(i).getDate());
                sheet.getRow(7 + i).getCell(2).setCellValue(list.get(i).getBuyNumber());
                sheet.getRow(7 + i).getCell(3).setCellValue(list.get(i).getContractNumber());
                sheet.getRow(7 + i).getCell(4).setCellValue(list.get(i).getDanNumber());
                sheet.getRow(7 + i).getCell(5).setCellValue(list.get(i).getProductName());
                sheet.getRow(7 + i).getCell(6).setCellValue(list.get(i).getGuiGe());
                sheet.getRow(7 + i).getCell(7).setCellValue(list.get(i).getCaiZhi());
                sheet.getRow(7 + i).getCell(8).setCellValue("");
                sheet.getRow(7 + i).getCell(9).setCellValue(list.get(i).getDanWei());
                sheet.getRow(7 + i).getCell(10).setCellValue(list.get(i).getNumber());
                sheet.getRow(7 + i).getCell(11).setCellValue(list.get(i).getOnePrice());
                sheet.getRow(7 + i).getCell(12).setCellValue(list.get(i).getPrice());
                sum += Integer.parseInt(list.get(i).getPrice());
                System.out.println("正在生成第" + i + "行数据，请稍等.......");
            }

            sheet.getRow(7 + list.size() + 2).getCell(3).setCellValue(sum);

            FileOutputStream out = null;
            out = new FileOutputStream(fileName);
            workbook.write(out);
            out.close();
            System.out.println("对账单生成成功");
        }
    }
    public static List<DemoData> getExcelValue() throws IOException {




       List<DemoData> list = new ArrayList<DemoData>();
       DemoData demoData = null;
        File file = new File("D:\\excelOption");
        File[] files = file.listFiles();
        int count = 1;
        for (int i = 0; i < files.length; i++) {
            FileInputStream excelFileInputStream = new FileInputStream(files[i]);
            XSSFWorkbook workbook = new XSSFWorkbook(excelFileInputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFRow row;
            for (int j = 9; j < sheet.getLastRowNum() - 2; j++) {
               demoData = new DemoData();

                row = sheet.getRow(j);
                if (!row.getCell(1).getStringCellValue().equals("")) {
                    for (int k = 1; k < row.getLastCellNum() - 2; k++) {
                        XSSFCell cell = row.getCell(k);
                        String cellValue = getCellValue(cell);

                       switch (k) {
                            case 1:
                                demoData.setProductName(cellValue);
                                break;
                            case 2:
                                demoData.setCaiZhi(cellValue);
                                break;
                            case 3:
                                demoData.setGuiGe(cellValue);
                                break;
                            case 4:
                                demoData.setDanWei(cellValue);
                                break;
                            case 5:
                                demoData.setNumber(cellValue);
                                break;
                            case 6:
                                demoData.setOnePrice(cellValue);
                                break;
                            case 7:
                                demoData.setPrice(cellValue);
                                break;
                            case 8:
                                demoData.setContractNumber(cellValue);
                                break;

                        }
                    }
                    demoData.setBuyNumber(getCellValue(sheet.getRow(7).getCell(8)).substring(3));
                    demoData.setDate(getCellValue(sheet.getRow(5).getCell(7)));
                    demoData.setDanNumber(getCellValue(sheet.getRow(6).getCell(7)));

                    demoData.setId(String.valueOf(count++));
                   list.add(demoData);
                } else {
                    break;
                }
            }

        }
        return  list;
    }


    public static String getCellValue(Cell cell) {
        String cellValue = "";
        // 以下是判断数据的类型
        switch (cell.getCellType()) {
            case NUMERIC: // 数字
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    cellValue = sdf.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(cell.getNumericCellValue())).toString();
                } else {
                    DataFormatter dataFormatter = new DataFormatter();
                    cellValue = dataFormatter.formatCellValue(cell);
                }
                break;
            case STRING: // 字符串
                cellValue = cell.getStringCellValue();
                break;
            case BOOLEAN: // Boolean
                cellValue = cell.getBooleanCellValue() + "";
                break;
            case FORMULA: // 公式
                cellValue = cell.getCellFormula() + "";
                break;
            case BLANK: // 空值
                cellValue = "";
                break;
            case ERROR: // 故障
                cellValue = "非法字符";
                break;
            default:
                cellValue = "未知类型";
                break;
        }
        return cellValue;
    }


    public static boolean writeXlsx(String fileName, int row, int column,

                                    String content) {


        boolean flag = false;

        FileOutputStream out = null;

        XSSFWorkbook xwb;

        try {

            xwb = new XSSFWorkbook(new FileInputStream(fileName));

            XSSFSheet xSheet = xwb.getSheetAt(0);

            XSSFCell xCell = xSheet.createRow(row).createCell(column);

            xCell.setCellValue(content);

            out = new FileOutputStream(fileName);

            xwb.write(out);

            out.close();

            flag = true;

        } catch (IOException e) {


            e.printStackTrace();

        } catch (RuntimeException e) {

            e.printStackTrace();

        }

        return flag;

    }


    public static void updateExcel(File exlFile, String sheetName, int col,

                                   int row, String value) throws Exception {

        FileInputStream fis = new FileInputStream(exlFile);

        HSSFWorkbook workbook = new HSSFWorkbook(fis);

// workbook.

        HSSFSheet sheet = workbook.getSheet(sheetName);

        HSSFCell mycell = sheet.createRow(row).createCell(col);

        mycell.setCellValue(value);

        HSSFRow r = sheet.getRow(row);

        HSSFCell cell = r.getCell(col);

// int type=cell.getCellType();

        String str1 = cell.getStringCellValue();

// 这里假设对应单元格原来的类型也是String类型

        cell.setCellValue(value);

        System.out.println("单元格原来值为" + str1);

        System.out.println("单元格值被更新为" + value);

        fis.close();// 关闭文件输入流

        FileOutputStream fos = new FileOutputStream(exlFile);

        workbook.write(fos);

        fos.close();// 关闭文件输出流

    }


    public static void simpleRead(String path) throws Exception {
        // 写法1：JDK8+ ,不用额外写一个DemoDataListener
        // since: 3.0.0-beta1
        // String fileName = TestFileUtil.getPath() + "demo" + File.separator + "demo.xlsx";
        File file = new File("D:\\excelOption");
        File[] files = file.listFiles();
        for (File fileName : files) {
            EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().headRowNumber(9).doRead();

            FileInputStream excelFileInputStream = new FileInputStream(fileName);
            XSSFWorkbook workbook = new XSSFWorkbook(excelFileInputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFRow row = sheet.getRow(7);
            String caigouhao = row.getCell(8).getStringCellValue();
            XSSFRow row2 = sheet.getRow(5);
            Date date = row2.getCell(7).getDateCellValue();
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
            String format = simpleDateFormat.format(date);
            System.out.println(format);

            XSSFRow row3 = sheet.getRow(6);
            String danhao = row3.getCell(7).getRawValue();
            System.out.println(danhao);
            System.out.println(caigouhao);


            workbook.close();
        }


        // String fileName = path+"风阀2套2022.1.24.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        // 这里每次会读取3000条数据 然后返回过来 直接调用使用数据就行
        //  EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().headRowNumber(9).doRead();

    }
}
