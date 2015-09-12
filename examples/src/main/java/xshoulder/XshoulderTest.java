package xshoulder;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.subtlelib.poi.api.sheet.SheetContext;
import org.subtlelib.poi.api.workbook.WorkbookContext;
import org.subtlelib.poi.impl.workbook.WorkbookContextFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

/**
 * Created by serv on 15/8/6.
 */
public class XshoulderTest {
    public static void main(String[] args) throws IOException {

        List<Map<Object,Object>> dataList = new ArrayList<Map<Object, Object>>();

        WorkbookContext workbook = WorkbookContextFactory.useWorkbook(new HSSFWorkbook(new FileInputStream("检察官学院信息.xls")));

        SheetContext sheetContext = workbook.useSheet("学生信息统计表");
//        SheetContext sheetContext = workbook.useSheet("学校信息");
        Sheet nativeSheet = sheetContext.getNativeSheet();

        Iterator<Row> iterator = nativeSheet.iterator();
        //skip the 0 row
        iterator.next();
        while (iterator.hasNext()){
            Row next = iterator.next();
            Iterator<Cell> cellIterator = next.cellIterator();
            Map<Object,Object> rowMap = new HashMap<Object, Object>();
            while (cellIterator.hasNext()){
                Cell cell = cellIterator.next();
                Object headValue = getValue(nativeSheet.getRow(0).getCell(cell.getColumnIndex()));
                Object columnValue = getValue(cell);
                rowMap.put(headValue,columnValue);
            }
            dataList.add(rowMap);
        }

        System.out.println(dataList.toString());
    }


    private static Object getValue(Cell cell){
        if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
            return cell.getNumericCellValue();
        }else if(cell.getCellType()==Cell.CELL_TYPE_BOOLEAN){
            return cell.getBooleanCellValue();
        }else if(cell.getCellType()==Cell.CELL_TYPE_ERROR){
            return cell.getErrorCellValue();
        }else if(cell.getCellType()==Cell.CELL_TYPE_FORMULA){
            return cell.getCellFormula();
        }else{
            return cell.getStringCellValue();
        }
    }
}
