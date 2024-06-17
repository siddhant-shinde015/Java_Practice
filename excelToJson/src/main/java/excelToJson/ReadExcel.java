package excelToJson;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class ReadExcel {

    public List<Map<String, Object>> readExcelFile(String filePath) {
        List<Map<String, Object>> dataList = new ArrayList<>();

        try {
        	FileInputStream fis = new FileInputStream(filePath);
        	Workbook workbook = null;
        	if(filePath.endsWith(".xlsx")) {
        		workbook = new XSSFWorkbook(fis);
        	} else if(filePath.endsWith(".xls")) {
        		workbook = new HSSFWorkbook(fis);
        	} else {
        		
        	}
          
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            List<String> headers = new ArrayList<>();      //keys
            
            if (rowIterator.hasNext()) {
                Row headerRow = rowIterator.next();
                for (Cell cell : headerRow) {
                    headers.add(cell.getStringCellValue());
                }
            }

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Map<String, Object> dataMap = new HashMap<>();
                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = row.getCell(i);                //cell at current column index
                    Object value = getCellValue(cell);			//retrieves the cell value
                    dataMap.put(headers.get(i), value);         //Puts the cell value in the map with the corresponding header as the key.
                }
                dataList.add(dataMap);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return dataList;
    }

    private Object getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return null;
        }
    }

    public static void main(String[] args) {
        ReadExcel excelReader = new ReadExcel();
        List<Map<String, Object>> dataList = excelReader.readExcelFile("C:\\Users\\hp\\Downloads\\Stock_Sample.xlsx");
//        System.out.println(dataList);
        for (Map<String, Object> dataMap : dataList) {
            System.out.println(dataMap);
        }
    }
}

