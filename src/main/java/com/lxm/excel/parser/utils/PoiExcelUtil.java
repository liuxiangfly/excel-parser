package com.lxm.excel.parser.utils;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class PoiExcelUtil {

    
    /** 
     * 创建新excel. 
     * @param fileDir  excel的路径 
     * @param sheetNames 要创建的表格索引 
     */  
    public static void createExcel(String fileDir, String... sheetNames){  
        //创建workbook  
        Workbook workbook = new XSSFWorkbook();
        //添加Worksheet（不添加sheet时生成的xls文件打开时会报错)
        if(sheetNames != null && sheetNames.length > 0){
            for(String sheetName: sheetNames){
                workbook.createSheet(sheetName);  
            }
        }else{
            workbook.createSheet("sheet1");  
        }
        //新建文件  
        FileOutputStream out = null;  
        try {  
            out = new FileOutputStream(fileDir);  
            workbook.write(out);
            out.close();
        } catch (Exception e) {  
            throw new RuntimeException(e);
        } 
    } 
    
    /** 
     * 创建新excel. 
     * @param fileDir  excel的路径 
     * @param sheetNames 要创建的表格索引 
     */  
    public static void createExcel(File file,String... sheetNames){  
        createExcel(file.getAbsolutePath(), sheetNames);
    } 


    /**
     * 解析excel，将结果转换为指定模型的列表
     * 
     * @param sheetFile
     * @param clazz
     * @param feilds
     * @param sheetIndex 读取第几个sheetIndex
     * @param spkit 跳过多少行
     * @param spkitEnd 最后几行不解析
     * @return
     */
    public static <T> List<T> parseExcelForModel(File sheetFile, Class<T> clazz, String[] feilds, int sheetIndex, int spkit, int spkitEnd) {
        List<T> result = new ArrayList<>();
        try {
            List<String[]> date = parseExcel(sheetFile, sheetIndex, spkit, spkitEnd);
            if (date.size() > 0) {
                try {
                    result = convertDataToDTO(clazz, date, feilds);
                } catch (Exception e) {
                    throw new RuntimeException(e);

                }
            }
            return result;
        } catch (Exception e) {
        	throw new RuntimeException(e);
        }
    }

    /**
     * 解析excel
     * 
     * @param sheetFile
     * @param sheetIndex
     * @param spkit
     * @param spkitEnd
     * @return
     */
    public static List<String[]> parseExcel(File sheetFile, int sheetIndex, int spkit, int spkitEnd){
        List<String[]> list = new ArrayList<String[]>();
        Row row = null;
        String value = null;
        Workbook wb;
        try {
            wb = WorkbookFactory.create(new FileInputStream(sheetFile));
        } catch (Exception e) {
            throw new RuntimeException(sheetFile.getAbsolutePath(), e);
        }
        Sheet sheet = wb.getSheetAt(sheetIndex);
        int rowNum = sheet.getLastRowNum();//总行数 
        if (rowNum == 0) {
            throw new RuntimeException(sheetFile.getAbsolutePath() + "，要读取的文件为空，请确认一遍！解析的文件为EXCEL文件。方法为：getExcelData");
        }
        if (spkit > rowNum + 1) {
            throw new RuntimeException(sheetFile.getAbsolutePath() + "，要读取的文件为空，请确认一遍！解析的文件为EXCEL文件。方法为：getExcelData");
        }
        if (spkitEnd < 0) {
            spkitEnd = 0;
        }

        for (int k = spkit; k <= rowNum - spkitEnd; k++) {
            try {
                row = sheet.getRow(k);
                if (null == row) {
                    break;
                }
                int columnNum = row.getLastCellNum();
                if(columnNum < 0){ // 该行无数据
                	continue;
                }
                String[] singleRow = new String[columnNum];
                int n = 0;
                for (int i = 0; i < columnNum; i++) {
                    Cell cell = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_BLANK:
                            singleRow[n] = "";
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            singleRow[n] = Boolean.toString(cell.getBooleanCellValue());
                            break;
                        //数值 
                        case Cell.CELL_TYPE_NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                singleRow[n] = formatDateString(cell.getDateCellValue());
                                // singleRow[n] = String.valueOf(cell.getDateCellValue());
                            } else {
                                cell.setCellType(Cell.CELL_TYPE_STRING);
                                String temp = cell.getStringCellValue();
                                if (StringUtils.isBlank(temp)) {
                                    temp = "";
                                }

                                // 则转换为BigDecimal类型的字符串 
                                singleRow[n] = temp.trim();
                            }
                            break;
                        case Cell.CELL_TYPE_STRING:
                            singleRow[n] = cell.getStringCellValue().trim();
                            break;
                        case Cell.CELL_TYPE_ERROR:
                            singleRow[n] = "";
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            //                            cell.setCellType(Cell.CELL_TYPE_STRING);
                            try {
                                singleRow[n] = cell.getStringCellValue();
                            } catch (IllegalStateException e) {
                                double v = cell.getNumericCellValue();
                                singleRow[n] = v + "";
                            }
                            if (value != null) {
                                singleRow[n] = value.replaceAll("#N/A", "").trim();
                            }
                            break;
                        default:
                            singleRow[n] = "";
                            break;
                    }
                    n++;
                }
                boolean isEmpty = true;
                for (String colunValue : singleRow) {
                    if (StringUtils.isNotBlank(colunValue)) {
                        isEmpty = false;
                        break;
                    }
                }
                if (isEmpty) {
                    continue;
                }//如果第一行为空，跳过 
                list.add(singleRow);

            } catch (Exception e) {
                String msg = sheetFile.getAbsolutePath() + "， 读取第" + k + "行数据出错(行数从0开始)";
                throw new RuntimeException(msg, e);
            }
        }
        return list;
    }
    
    private static byte[] getBytes(File file) {
    	try(InputStream in = new FileInputStream(file);
    			ByteArrayOutputStream buffer = new ByteArrayOutputStream();){
    		int len;
            byte[] data = new byte[100000];
            while ((len = in.read(data, 0, data.length)) != -1) {
                buffer.write(data, 0, len);
            }
            buffer.flush();
            return buffer.toByteArray();
    	}catch (Exception e) {
			throw new RuntimeException(e);
		}
    }

    /**
     * 导出excel
     * 
     * @param sheetFile
     * @param list
     * @param feilds
     * @param sheetIndex 第几个sheet
     * @param spkit 跳过第几行
     */
    public static <T> void exportExcel(File sheetFile, List<T> list, String[] feilds, int sheetIndex, int spkit) {
        T obj;
        int listLen = list.size();
        List<String[]> temp = new ArrayList<String[]>();
        int length = feilds.length;
        try {
            for (int i = 0; i < listLen; i++) {
                obj = list.get(i);
                for (int j = 0; j < length; j++) {
                    String[] arr = new String[length];
                    Field field = obj.getClass().getDeclaredField(feilds[j]);
                    field.setAccessible(true);
                    Object val = field.get(obj);
                    if ("java.util.Date".equals(field.getType().getName())) {
                        arr[j] = formatDateString((Date) val);
                    } else {
                        arr[j] = String.valueOf(val);
                    }
                    temp.add(arr);
                }
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        exportExcel(sheetFile, temp, sheetIndex, spkit);
    }
    
    /**
     * 导出excel
     * 
     * @param sheetFile
     * @param date
     * @param sheetIndex
     * @param spkit
     */
    public static void exportExcel(File sheetFile, List<String[]> date, int sheetIndex, int spkit){
        int listLen = date.size();
        //写入excel 
        Workbook wb;
        try {
        	wb = WorkbookFactory.create(new FileInputStream(sheetFile));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        Sheet sheet = wb.getSheetAt(sheetIndex);
        Row row = null;
        for (int i = 0; i < listLen; i++) {
            String[] arr = date.get(i);
            row = sheet.getRow(spkit + i);
            if(row == null){
                row = sheet.createRow(spkit + i);
            }
            for (int j = 0; j < arr.length; j++) {
                Cell cell = row.getCell(j);
                if (cell == null) {
                    cell = row.createCell(j);
                }
                cell.setCellValue(arr[j]);
            }
        }
        try{
            FileOutputStream out = new FileOutputStream(sheetFile);
            wb.write(out);
            out.close();
        }catch(Exception e){
            throw new RuntimeException(e);
        }
    }

    /**
     * 将excel列号转换为数字下标，如A->0,B->1
     * @param column
     * @return
     */
    public static int column2Index(String column){
    	int index = 0;
    	int pow = 1; // 初始为26的0次方，后续每循环一次乘以26一次
    	for(int i = column.length() - 1; i >= 0; i--){
    		index += (Character.toUpperCase(column.charAt(i)) - 64) *  pow;
    		pow *= 26;
    	}
    	return index - 1;
    }

    private static <T> List<T> convertDataToDTO(Class<T> clazz, List<String[]> dataList, String[] fieldNames)
            throws InstantiationException, IllegalAccessException {
        List<T> temp = new ArrayList<T>();
        int dataLength = dataList.size();
        T obj = null;
        for (int i = 0; i < dataLength; i++) {
            try {
                obj = clazz.newInstance();
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
            String[] data = dataList.get(i);
            for (int j = 0; j < fieldNames.length; j++) {
                try {
                    if(j > data.length - 1)/* 属性下标超过数据最大下标 */{
                        break;
                    }
                    Field field = clazz.getDeclaredField(fieldNames[j]);
                    field.setAccessible(true);
                    if ("java.util.Date".equals(field.getType().getName())) {
                        field.set(obj, StringUtils.isBlank(data[j]) ? null : parseString(data[j].trim()));
                    } else if ("java.lang.String".equals(field.getType().getName())) {
                        field.set(obj, StringUtils.isBlank(data[j]) ? null : data[j].trim());
                    } else if ("java.lang.Long".equals(field.getType().getName())) {
                        field.set(obj, StringUtils.isBlank(data[j]) ? null : Long.valueOf(data[j].trim()));
                    } else if ("java.lang.Integer".equals(field.getType().getName())) {
                        field.set(obj, StringUtils.isBlank(data[j]) ? null : Integer.valueOf(data[j].trim()));
                    } else if ("java.lang.Double".equals(field.getType().getName())) {
                        field.set(obj, StringUtils.isBlank(data[j]) ? null : Double.valueOf(data[j].trim()));
                    } else if ("java.lang.Float".equals(field.getType().getName())) {
                        field.set(obj, StringUtils.isBlank(data[j]) ? null : Float.valueOf(data[j].trim()));
                    } else if ("java.math.BigDecimal".equals(field.getType().getName())) {
                        field.set(obj, StringUtils.isBlank(data[j]) ? null : new BigDecimal(data[j].trim()));
                    } else {
                        throw new RuntimeException("目前暂时还不支持" + field.getType().getName());
                    }
                } catch (Exception e) {
                    int row = i + 1;
                    int col = j + 1;
                    throw new RuntimeException("第" + row + "行" + col + "列数格式错误;", e);
                }
            }
            temp.add(obj);
        }
        return temp;
    }

    private static Date parseString(String date) {
        try {
            return org.apache.commons.lang.time.DateUtils.parseDate(date, new String[] { "yyyyMMddHHmmss",
                    "yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd HH:mm:ss", "yyyy-MM-dd", "yyyy/MM/dd", "yy-MM-dd HH:mm:ss",
                    "yy/MM/dd hh:mm:ss", "yyyyMMdd" });
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 功能: 将字符串转换为指定格式的日期返回
     * @param date (yyyy-MM-dd HH:mm:ss)
     * @return
     */
    private static String formatDateString(Date date) {
        return formatDateString(date, "yyyy-MM-dd HH:mm:ss");
    }

    /**
     * 功能: 将字符串转换为指定格式的日期返回
     * @param date
     * @param formatString
     * @return
     */
    private static String formatDateString(Date date, String formatString) {
        if (null == date)
            return null;
        SimpleDateFormat formatter;
        formatter = new SimpleDateFormat(formatString);
        String ctime = formatter.format(date);
        return ctime;
    }
    
}
