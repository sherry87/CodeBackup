package com.execl;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 *                 @描述：测试excel读取
 * 
 *               导入的jar包
 * 
 *               poi-3.8-beta3-20110606.jar
 * 
 *               poi-ooxml-3.8-beta3-20110606.jar
 * 
 *               poi-examples-3.8-beta3-20110606.jar
 * 
 *               poi-excelant-3.8-beta3-20110606.jar
 * 
 *               poi-ooxml-schemas-3.8-beta3-20110606.jar
 * 
 *               poi-scratchpad-3.8-beta3-20110606.jar
 * 
 *               xmlbeans-2.3.0.jar
 * 
 *               dom4j-1.6.1.jar
 * 
 *               jar包官网下载地址：http://poi.apache.org/download.html
 * 
 *               下载poi-bin-3.8-beta3-20110606.zipp
 * 
 */

public class ImportExecl {
    /** 总行数 */
    private int totalRows = 0;
    /** 总列数 */
    private int totalCells = 0;
    /** 错误信息 */
    private String errorInfo;
    /** 构造方法 */
    public ImportExecl() {
    }


    public int getTotalRows() {
        return totalRows;
    }
    public int getTotalCells() {
        return totalCells;
    }
    public String getErrorInfo() {
        return errorInfo;
    }

    public boolean validateExcel(String filePath) {
        /** 检查文件名是否为空或者是否是Excel格式的文件 */
        if (filePath == null
                || !(WDWUtil.isExcel2003(filePath) || WDWUtil
                        .isExcel2007(filePath))) {
            errorInfo = "文件名不是excel格式";
            return false;
        }
        /** 检查文件是否存在 */
        File file = new File(filePath);
        if (file == null || !file.exists()) {
            errorInfo = "文件不存在";
            return false;
        }
        return true;
    }

    public List<List<List<String>>> read(String filePath) {
    	List<List<List<String>>> sheetLst = new ArrayList<List<List<String>>>();
        InputStream is = null;
        try {
            /** 验证文件是否合法 */
            if (!validateExcel(filePath)) {
                System.out.println(errorInfo);
                return null;
            }
            /** 判断文件的类型，是2003还是2007 */
            boolean isExcel2003 = true;
            if (WDWUtil.isExcel2007(filePath)) {
                isExcel2003 = false;
            }
            /** 调用本类提供的根据流读取的方法 */
            File file = new File(filePath);
            is = new FileInputStream(file);
            sheetLst = read(is, isExcel2003);
            is.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    is = null;
                    e.printStackTrace();
                }
            }
        }
        /** 返回最后读取的结果 */
        return sheetLst;
    }


    public List<List<List<String>>> read(InputStream inputStream, boolean isExcel2003) {
    	List<List<List<String>>> sheetLst = new ArrayList<List<List<String>>>();
        try {
            /** 根据版本选择创建Workbook的方式 */
            Workbook wb = null;
            if (isExcel2003) {
                wb = new HSSFWorkbook(inputStream);
            } else {
                wb = new XSSFWorkbook(inputStream);
            }
            sheetLst = read(wb);
        } catch (IOException e) {
        
            e.printStackTrace();
        }
        return sheetLst;
    }
    

    private List<List<List<String>>> read(Workbook wb) {
        List<List<List<String>>> sheetLst = new ArrayList<List<List<String>>>();
        int sheets = wb.getNumberOfSheets();
        for(int i=0;i<sheets;i++){
        	List<List<String>> dataLst = new ArrayList<List<String>>();
        	/** 得到第一个shell */
            Sheet sheet = wb.getSheetAt(i);
            /** 得到Excel的行数 */
            this.totalRows = sheet.getPhysicalNumberOfRows();
            /** 得到Excel的列数 */
            if (this.totalRows >= 1 && sheet.getRow(0) != null) {
                this.totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
            }
            /** 循环Excel的行 */
            for (int r = 0; r < this.totalRows; r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }
                List<String> rowLst = new ArrayList<String>();
                /** 循环Excel的列 */
                for (int c = 0; c < this.getTotalCells(); c++) {
                    Cell cell = row.getCell(c);
                    String cellValue = "";
                    if (null != cell) {
                        // 以下是判断数据的类型
                        switch (cell.getCellType()) {
                        case HSSFCell.CELL_TYPE_NUMERIC: // 数字
                        	if(cell.getCellStyle().getDataFormat()==31||cell.getCellStyle().getDataFormat()==176){
                                SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日");  
                                Date date = cell.getDateCellValue();  
                                cellValue = sdf.format(date);
                                String m = cellValue.substring(5,7);
                				String d = cellValue.substring(8,10);
                				cellValue=cellValue.substring(0,5)+Integer.valueOf(m)+"月"+Integer.valueOf(d)+"日";
                            } else {  
                            	 cellValue = cell.getNumericCellValue() + "";
                            } 
                            break;
                        case HSSFCell.CELL_TYPE_STRING: // 字符串
                            cellValue = cell.getStringCellValue();
                            break;
                        case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
                            cellValue = cell.getBooleanCellValue() + "";
                            break;
                        case HSSFCell.CELL_TYPE_FORMULA: // 公式
                            cellValue = cell.getCellFormula() + "";
                            break;
                        case HSSFCell.CELL_TYPE_BLANK: // 空值
                            cellValue = "";
                            break;
                        case HSSFCell.CELL_TYPE_ERROR: // 故障
                            cellValue = "非法字符";
                            break;
                        default:
                            cellValue = "未知类型";
                            break;
                        }
                        
                       
                    }
                    rowLst.add(cellValue);
                }
                /** 保存第r行的第c列 */
                dataLst.add(rowLst);
            }
            
            sheetLst.add(dataLst);
        }
        
        return sheetLst;
    }
    public static void main(String[] args) throws Exception {
        ImportExecl poi = new ImportExecl();
        // List<List<String>> list = poi.read("d:/aaa.xls");
        List<List<List<String>>> list = poi.read("E:\\1.xlsx");
        /*if (list != null) {
            for (int i = 0; i < list.size(); i++) {
                System.out.println("第" + (i) + "sheet");
                List<List<String>> dataList = list.get(i);
                for (int j = 0; j < dataList.size(); j++) {
                	 System.out.println("第" + (j) + "行");
                	 
                	 List<String> cellList = dataList.get(j);
                	 for (int k = 0; k < cellList.size(); k++) {
                		 System.out.print("    " + cellList.get(k));
                	 }
                	 System.out.println();
                }
            }

        }*/

    }

}

class WDWUtil {
    public static boolean isExcel2003(String filePath) {
        return filePath.matches("^.+\\.(?i)(xls)$");
    }
    public static boolean isExcel2007(String filePath) {
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }
}