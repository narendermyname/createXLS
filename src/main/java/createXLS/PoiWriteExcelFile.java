package createXLS;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class PoiWriteExcelFile {

	public static void main(String[] args) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {

		//createNormalXls();
		
		String header = "EMP_NAME,EMP_AGE";
		List<String> empToStringList = callMethodDynamic(Arrays.asList(new Employee[]{
				new Employee("Narender",29,"Narendermyname"),
				new Employee("Narender2",29,"Narendermyname2"),
				new Employee("Narender3",29,"Narendermyname3")
		}),header.replaceAll("_","").toLowerCase());
		System.out.println(empToStringList);
		createXlsFromObject(empToStringList,header);
	}

	private static List<String> callMethodDynamic(List<Employee> emps,String headers) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException{
		List<String> methods = getObjMethods(Employee.class);
		List<String> empToStringList =new ArrayList<String>();
		
		for(Employee emp:emps){
			String row ="";
			for(String methodName :methods){
				if(headers.contains(methodName.replace("get","").toLowerCase())){
					try {
						Object value = emp.getClass().getMethod(methodName).invoke(emp, null);
						row += (row == "" ?value:","+value);
					} catch (SecurityException |NoSuchMethodException e) {
						System.out.println(e);
					}
				}
			}
			empToStringList.add(row);
		}

		return empToStringList;
	}

	private static <T> List<String> getObjMethods(Class<T> classA){

		return Arrays.asList(classA.getMethods()).stream().filter(method -> method.getName().contains("get")).map(method -> method.getName()).collect(Collectors.toList());

	}

	private static void createNormalXls() {
		try {
			FileOutputStream fileOut = new FileOutputStream("poi-test.xls");
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet worksheet = workbook.createSheet("POI Worksheet");

			// index from 0,0... cell A1 is cell(0,0)
			HSSFRow row1 = worksheet.createRow((short) 0);

			HSSFCell cellA1 = row1.createCell((short) 0);
			cellA1.setCellValue("Hello");
			HSSFCellStyle cellStyle = workbook.createCellStyle();
			cellStyle.setFillForegroundColor(HSSFColor.GOLD.index);
			cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			cellA1.setCellStyle(cellStyle);

			HSSFCell cellB1 = row1.createCell((short) 1);
			cellB1.setCellValue("Goodbye");
			cellStyle = workbook.createCellStyle();
			cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
			cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			cellB1.setCellStyle(cellStyle);

			HSSFCell cellC1 = row1.createCell((short) 2);
			cellC1.setCellValue(true);

			HSSFCell cellD1 = row1.createCell((short) 3);
			cellD1.setCellValue(new Date());
			cellStyle = workbook.createCellStyle();
			cellStyle.setDataFormat(HSSFDataFormat
					.getBuiltinFormat("m/d/yy h:mm"));
			cellD1.setCellStyle(cellStyle);

			workbook.write(fileOut);
			fileOut.flush();
			fileOut.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	private static void createXlsFromObject(List<String> emps,String headers){
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Sample sheet");
		
		sheet.createFreezePane(0,1);
		
		Map<String, Object[]> data = new HashMap<String, Object[]>();
		
		Object headerObj[] = headers.split(",");
		data.put("1",headerObj);
		for(int i = 1;i <= emps.size();i++){
			Object headerObjA[] = emps.get(i-1).split(",");
			data.put(String.valueOf(i+1), headerObjA);
		}
		
		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object [] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if(obj instanceof Date) 
					cell.setCellValue((Date)obj);
				else if(obj instanceof Boolean)
					cell.setCellValue((Boolean)obj);
				else if(obj instanceof String)
					cell.setCellValue((String)obj);
				else if(obj instanceof Double)
					cell.setCellValue((Double)obj);
			}
		}

		try {
			FileOutputStream out = 
					new FileOutputStream(new File("new.xls"));
			workbook.write(out);
			out.close();
			System.out.println("Excel written successfully..");

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}

class Employee{

	private String empName;
	private int empAge;
	private String empEmail;
	public String getEmpName() {
		return empName;
	}
	public void setEmpName(String empName) {
		this.empName = empName;
	}
	public int getEmpAge() {
		return empAge;
	}
	public void setEmpAge(int empAge) {
		this.empAge = empAge;
	}
	public String getEmpEmail() {
		return empEmail;
	}
	public void setEmpEmail(String empEmail) {
		this.empEmail = empEmail;
	}
	public Employee(String empName, int empAge, String empEmail) {
		super();
		this.empName = empName;
		this.empAge = empAge;
		this.empEmail = empEmail;
	}
	@Override
	public String toString() {
		return "Employee [empName=" + empName + ", empAge=" + empAge + ", empEmail=" + empEmail + "]";
	}
}