package createXLS;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

public class PoiWriteExcelFile {

	private static final String GET ="get";
	//private static final String UNDER_SCORE ="_";
	private static final String COMMA = ",";
	private static final String EMPTY ="";
	private static final String DATE_FORMATE ="MM/dd/yyyy";
	private static final String DATA_FORMATE = "0.00";
	
	
	public static void main(String[] args) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, ParseException {

		//createNormalXls();

		String header = "STREET_TWO,EMP_AGE,EMP_NAME,STREET_ONE,DOB,EXPENCESS,SALARY";
		//List<String> empToStringList = callMethodDynamic(,header.replaceAll(UNDER_SCORE,"").toLowerCase());
		//System.out.println(empToStringList);
		createXlsFromObject(Arrays.asList(new Employee[]{
				new Employee("Narender",20,"Narendermyname",new Date(),Long.valueOf(1200),Double.valueOf(12399.0008),new Address("Add A1 SDS VS  SDDDDFFFFFFFFFFF","Add B1")),
				new Employee("Narender2",22,"Narendermyname2",new Date(),Long.valueOf(1200),Double.valueOf(12399.0228),new Address("Add A2","Add B2")),
				new Employee("Narender3",23,"Narendermyname3",new Date(),Long.valueOf(1200),Double.valueOf(12399.0238),new Address("Add A3","Add B3"))
		}),header);
	}

	@SuppressWarnings("unused")
	private static List<String> callMethodDynamic(List<Employee> emps,String headers) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException{
		Map<String,String> parentMethods = getObjMethods(Employee.class);
		Map<String,String> childMethods = getObjMethods(Address.class);
		List<String> empToStringList =new ArrayList<String>();
		for(Employee emp:emps){
			String row =EMPTY;
			for(String header :headers.split(COMMA)){
				if(parentMethods.containsKey(header)){
					row += (row == EMPTY ?getValue(emp,parentMethods.get(header)):COMMA+getValue(emp,parentMethods.get(header)));
				}
				if(emp.getAdd() != null && childMethods.containsKey(header)){
					row += (row == EMPTY ?getValue(emp.getAdd(),childMethods.get(header)):COMMA+getValue(emp.getAdd(),childMethods.get(header)));
				}
			}
			empToStringList.add(row);
		}
		return empToStringList;
	}
	private static Object getValue(Object obj,String methodName){
		Object value = null;
		try {
			value = obj.getClass().getMethod(methodName).invoke(obj, null);

		} catch (SecurityException |NoSuchMethodException | IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
			System.out.println(e);
		}

		return value;
	}

	private static <T> Map<String, String> getObjMethods(Class<T> classA){

		return Arrays.asList(classA.getMethods()).stream().filter(method -> method.getName().contains(GET)).collect(
				Collectors.toMap(m -> m.getName().replace(GET,EMPTY).toLowerCase().toString(),method -> method.getName() )
				);

	}

	@SuppressWarnings({ "unused", "deprecation" })
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

	private static void createXlsFromObject(List<Employee> emps,String headers) throws ParseException{
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Sample sheet");

		sheet.createFreezePane(0,1);

		int rownum = 0;
		int cellnum = 0;
		Row row = sheet.createRow(rownum++);
		for (String head : headers.split(COMMA)) {
			Cell cell = row.createCell(cellnum++);
			cell.setCellValue(head);
		}
			
		for (Employee emp : emps) {
			row = sheet.createRow(rownum++);
			cellnum = 0;
			for (String key : headers.split(COMMA)) {
				Object obj =emp.getFieldValue(key);
				Cell cell = row.createCell(cellnum++);
				if(obj instanceof Date){
					SimpleDateFormat datetemp = new SimpleDateFormat(DATE_FORMATE);
					Date cellValue = datetemp.parse(datetemp.format(obj));
					cell.setCellValue(cellValue);

					//binds the style you need to the cell.
					HSSFCellStyle dateCellStyle = workbook.createCellStyle();
					short df = workbook.createDataFormat().getFormat("dd-mmm");
					dateCellStyle.setDataFormat(df);
					cell.setCellStyle(dateCellStyle);
				}
				else if(obj instanceof Long)
					cell.setCellValue((Long)obj);
				else if(obj instanceof Boolean)
					cell.setCellValue((Boolean)obj);
				else if(obj instanceof Integer)
					cell.setCellValue((Integer)obj);
				else if(obj instanceof Double){
					
					// Do this only once per file
					CellStyle cellStyle = workbook.createCellStyle();
					cellStyle.setDataFormat(
					    workbook.getCreationHelper().createDataFormat().getFormat(DATA_FORMATE));
					cell.setCellValue((Double)obj);
					cell.setCellStyle(cellStyle);
				}else if(obj instanceof String)
					cell.setCellValue((String)obj);
					
				//sheet.autoSizeColumn(cellnum);
			}
		}
		for(int i = 1;i <= emps.size();i++){
			sheet.autoSizeColumn(i);
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
	private Address add;
	private Date dob;
	private Long expencess;
	private Double salary;

	public Address getAdd() {
		return add;
	}
	public void setAdd(Address add) {
		this.add = add;
	}
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
	
	
	public Date getDob() {
		return dob;
	}
	public void setDob(Date dob) {
		this.dob = dob;
	}
	
	public Long getExpencess() {
		return expencess;
	}
	public void setExpencess(Long expencess) {
		this.expencess = expencess;
	}
	public Double getSalary() {
		return salary;
	}
	public void setSalary(Double salary) {
		this.salary = salary;
	}
	public Employee(String empName, int empAge, String empEmail,Date dob,Long expencess,Double salary,Address add) {
		super();
		this.empName = empName;
		this.empAge = empAge;
		this.empEmail = empEmail;
		this.add = add;
		this.dob=dob;
		this.expencess = expencess;
		this.salary = salary;
	}
	@Override
	public String toString() {
		return "Employee [empName=" + empName + ", empAge=" + empAge + ", empEmail=" + empEmail + "]";
	}
	public Object getFieldValue(String name){
		switch(name){
		case"EMP_NAME":return getEmpName();
		case "EMP_AGE":return getEmpAge();
		case "DOB":return getDob();
		case "STREET_ONE":return getAdd().getStreetOne();
		case "STREET_TWO":return getAdd().getStreetTwo();
		case "EXPENCESS":return getExpencess();
		case "SALARY":return getSalary();
		}
		return null;
	}
}

class Address{
	private String streetOne;
	private String StreetTwo;
	public String getStreetOne() {
		return streetOne;
	}
	public void setStreetOne(String streetOne) {
		this.streetOne = streetOne;
	}
	public String getStreetTwo() {
		return StreetTwo;
	}
	public void setStreetTwo(String streetTwo) {
		StreetTwo = streetTwo;
	}
	public Address(String streetOne, String streetTwo) {
		super();
		this.streetOne = streetOne;
		StreetTwo = streetTwo;
	}



}
