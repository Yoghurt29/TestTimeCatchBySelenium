package com.yo.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.UUID;
import java.util.function.Function;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.junit.Test;

import com.fasterxml.jackson.databind.ObjectMapper;




/**
 * <xmp>
 * Excel 到List<Bean>（List<T>） 或List<Map>读取工具
 * List<Map>使用JSON工具转化成List<Bean>
 * 依赖poi jackson 包，MulitPartFiel操作 方法依赖JavaWeb容器
 * </xmp>
 * @author Trulon_Chu
 */
@SuppressWarnings(value = { "rawtypes","unchecked"})
public class ExcelUtil {
	/************************************典型用法demo********************************/
	/*******
	 * <xmp>
	 * 典型用法demo,请进入源码查看该方法
	 * 一个人对象的一个成员是地址对象，通过一个excel直接读出一个人对象的等效map
	 * 父Map包含name sex addess，Code等字段
	 * 其中addess字段的值也是Map ，包含addressCode，addressDetail字段
	 * 最后的父map是可以用JSON工具类直接转换成对应的类对象，List<map>则可直接转换成List<people>
	 * </xmp>
	 */
	//@Test
	public void test() throws Exception {
		File file=new File("D:/Users/Trulon_Chu/Desktop/People.xlsx");//excel的文件表头为 name sex addressCode addressDetail
		List<Map<String,Object>> list = ExcelUtil.readListFromExcel(file,new Function<ExcelUtil.CoverterContext, Object>(){
			@Override
			public Object apply(CoverterContext t) {
				switch (t.getCurrentField()) {
				case "name":
					//字段name啥规则也不做，直接返回即可
					return t.getCurrentValue();
				case "sex":
					//字段sex在excel中写的是男/女，而实体类，或说数据库存的是 man/woman，因此我们对该字段写一个转换规则如下
					if("男".equals(t.getCurrentValue())){
						return "Man";
					}else{
						return "Woman";
					}
				case "addressCode":
					//字段addressCode放在子addess上，同时也放在父map上
					t.putFieldToObject("addess", "addressCode", t.getCurrentValue());//putFieldToObject 该方法可创建子map，并放入属性和值
					//CONVERT_RESULT_TYPE_RETURNED 表示有返回值，有返回值则将加入当前map
					t.setConvertResultType(ExcelUtil.CoverterContext.CONVERT_RESULT_TYPE_RETURNED);
					return t.getCurrentValue();
				case "addressDetail":
					//字段addressDetail仅放在子map addess上
					t.putFieldToObject("addess", "addressDetail", t.getCurrentValue());
					//CONVERT_RESULT_TYPE_RETURNED 表示无返回值，有返回值则将不会加入当前map，对于地址详情来说，只存在于people的address上
					t.setConvertResultType(ExcelUtil.CoverterContext.CONVERT_RESULT_TYPE_NON_RETURNED);
					break;
				default:
					//其他字段也不做修改，直接返回
					return t.getCurrentValue();
				}
				return null;
			}});
		for (Map<String, Object> map : list) {
			System.out.println(map.toString());
		}
		//{addressCode=SH600, addess={addressCode=SH600, addressDetail=上海市600号}, sex=Man, name=tom}
		//{addressCode=SH601, addess={addressCode=SH601, addressDetail=上海市601号}, sex=Woman, name=Jerry}

	}
	/************************************导入工具********************************/
	/**
	 * 最大读取Excel文件列数100
	 */
	public static int MAX_FIELDS_COUNT=100;
	/**
	 * MuiltPartFileToDisk默认存储位置
	 * WEB容器下的指定临时目录req.getServletContext().getRealPath("/")+TEMP_EXCEL_FOLDER
	 */
	public static String TEMP_EXCEL_FOLDER="TempExcelFolder";
	/**
	 * 转换上下文
	 * 用于控制转换规则，用户自定转换规则时候回调函数的传参等[要点]
	 * @author Trulon_Chu
	 *
	 */
	public static class CoverterContext{//该类用于用户自定转换规则时候回调函数的传参[要点]
		/**
		 * 自定义转换规则时的结果类型，采用回调函数形式处理，表示返回了字段，正常put处理
		 */
		public static String CONVERT_RESULT_TYPE_RETURNED="returned";//自定义转换规则时的结果类型，采用回调函数形式处理，表示返回了字段，正常put处理
		/**
		 * 自定义转换规则时的结果类型，采用回调函数形式处理，表示已经在回调函数中处理了，将不会在put到当前map
		 */
		public static String CONVERT_RESULT_TYPE_NON_RETURNED="non_returned";//自定义转换规则时的结果类型，采用回调函数形式处理，表示已经在回调函数中处理了
		/**
		 * 转换到的对象map 或T
		 */
		private Object toObject; //转换到的对象map 或T
		/**
		 * 当前正在转换的字段
		 */
		private String currentField; //当前正在转换的字段
		/**
		 * 当前正在转换的字段的单元格值
		 */
		private String currentValue; //当前正在转换的字段的单元格值
		private String convertResultType=CONVERT_RESULT_TYPE_RETURNED;
		public CoverterContext() {
			super();
		}
		public Map getToObject() {
			return (Map) toObject;
		}
		public void setToObject(Object toObject) {
			this.toObject = toObject;
		}
		public String getCurrentField() {
			return currentField;
		}
		public void setCurrentField(String currentField) {
			this.currentField = currentField;
		}
		public String getCurrentValue() {
			return currentValue;
		}
		public void setCurrentValue(String currentValue) {
			this.currentValue = currentValue;
		}
		public String getConvertResultType() {
			return convertResultType;
		}
		public void setConvertResultType(String convertResultType) {
			this.convertResultType = convertResultType;
		}
		public CoverterContext(Object toObject, String currentField, String currentValue) {
			super();
			this.toObject = toObject;
			this.currentField = currentField;
			this.currentValue = currentValue;
		}
		/**
		 * 根据所给表达式定位字段，若字段不存在则添加该字段（一个map），并将在该map中置入key为addField值为value的实体
		 * @param fieldExp 表达式定位，将在该处插入key 为 fieldExp 值为空Map 的实体
		 * @param addField 将在插入的map中插入的实体的key
		 * @param value 将在插入的map中插入的实体的value
		 */
		public void putFieldToObject(String fieldExp,String addField, Object value){
			this.setConvertResultType(CONVERT_RESULT_TYPE_NON_RETURNED);
			try {
				ExcelUtil.putFieldToObject(this.toObject, fieldExp, addField, value);
			} catch (Exception e) {
				//由于回调方法无法往上抛出异常，避免用户反复处理异常，在此处理[重要]
				//字段插入失败将不会中断，注意
				//参考Java编程思想 可将检查异常等转换为运行异常抛出，且保留原始异常栈轨迹信息[20170724]
				e.printStackTrace();
				throw new RuntimeException(e);
			}
		}
	}
	public ExcelUtil() {
		super();
	}
	/**
	 * [重载]按给定索引与值的Map修改excel文件对应列表头，并选取excel表头字段，读取excel指定列为元素为对象的List集合
	 */
	public static <T> List<T>  readListFromExcel(File fileExcel,Class<T> cls,Map<Integer,String> tableHeadAlterConfig,String[] fields ) throws Exception{
		return ExcelUtil.readListFromExcel(fileExcel, cls,null,tableHeadAlterConfig,fields);
	}
	/**
	 * [重载]选取excel列头字段，读取excel指定列为元素为对象的List集合
	 */
	public static <T> List<T>  readListFromExcel(File fileExcel,Class<T> cls,String[] fields ) throws Exception{
		return ExcelUtil.readListFromExcel(fileExcel, cls,null,null,fields);
	}
	/**
	 * [重载]通过选取excel列头字段，并自定义转换规则，读取excel指定列为元素为对象的List集合（也可在自定转换规则中选取字段，则fields传递null）
	 */
	public static <T> List<T>  readListFromExcel(File fileExcel,Class<T> cls,Function<CoverterContext, Object> fieldConverter,String[] fields ) throws Exception{
		return ExcelUtil.readListFromExcel(fileExcel, cls,fieldConverter,null,fields);
	}
	/**
	 * [重载]按给定索引与值的Map修改excel文件对应列表头，再按自定转换规则，读取excel指定列为元素为对象的List集合
	 */
	public static <T> List<T>  readListFromExcel(File fileExcel,Class<T> cls,Function<CoverterContext, Object> fieldConverter,Map<Integer,String> tableHeadAlterConfig) throws Exception{
		return ExcelUtil.readListFromExcel(fileExcel, cls,fieldConverter,tableHeadAlterConfig,null);
	}
	/**
	 * [重载]通过自定义转换规则，将从excel文件读取元素为对象的List集合
	 * @param fileExcel 文件
	 * @param cls 指定类class，实现自动封装到对象
	 * @param <xmp>fieldConverter 自定转换规则，可读取map的值为新的map，如student.course.name student对象的course为引用类型，可通过定义转换器实现构造出List<Map<studentField,Map<courseField,stringValue>>></xmp>
	 * @return T类型的对象的的List集合
	 * @throws Exception Exce单元格得String值不符合其对应的T的字段类型所需的格式，Excel单元格格式无法处理
	 */
	public static <T> List<T>  readListFromExcel(File fileExcel,Class<T> cls,Function<CoverterContext, Object> fieldConverter) throws Exception{
		return ExcelUtil.readListFromExcel(fileExcel, cls,fieldConverter,null,null);
	}
	/**
	 * [重载]通过自定转换规则，读取excel为对象集合
	 * @see ExcelInportUtil.readListFromExcel(File, Class<T>, Function<CoverterContext, Object>, Map<Integer, String>, String[])
	 * @param file 文件
	 * @param <xmp>fieldConverter 自定转换规则，可读取map的值为新的map，如student.course.name student对象的course为引用类型，可通过定义转换器实现构造出List<Map<studentField,Map<courseField,stringValue>>></xmp>
	 * @return Map类型的元素的List集合
	 * @throws Exception Excel单元格格式无法处理
	 */
	public static List<Map<String,Object>> readListFromExcel(File file,Function<CoverterContext, Object> fieldConverter) throws Exception{
		return ExcelUtil.readListFromExcel(file, fieldConverter, null);
	}
	/**
	 * [重载]选取excel列头字段，读取excel指定列为元素为Map的List集合
	 * @see ExcelInportUtil.readListFromExcel(File, Class<T>, Function<CoverterContext, Object>, Map<Integer, String>, String[])
	 * @param file 文件
	 * @param fields 指定所需的excel表头字段，传入null则转换所有字段
	 * @return Map类型的元素的List集合
	 * @throws Exception Excel单元格格式无法处理
	 */
	public static List<Map<String,Object>> readListFromExcel(File file,String[] fields) throws Exception{
		Function<CoverterContext, Object> fieldConverter=null;
		return ExcelUtil.readListFromExcel(file, fieldConverter, fields);
	}
	/**
	 * [重载]按给定索引与值的Map修改excel文件对应列表头，并选取excel表头字段，读取excel指定列为元素为Map的List集合
	 * @see ExcelInportUtil.readListFromExcel(File, Class<T>, Function<CoverterContext, Object>, Map<Integer, String>, String[])
	 * @param file 文件
	 * @param tableHeadAlterConfig 按给定索引与值修改excel文件表头，再读取excel 
	 * @param fields 指定所需的excel表头字段，传入null则转换所有字段
	 * @return K表头字段 V单元格值构成的Map 置入ArrayList
	 * @throws Exception Excel单元格格式无法处理
	 */
	public static List<Map<String,Object>> readListFromExcel(File file,Map<Integer,String> tableHeadAlterConfig,String[] fields) throws Exception{
		Function<CoverterContext, Object> fieldConverter=null;
		if(null!=tableHeadAlterConfig){
			alterTableHead(file,tableHeadAlterConfig);
		}
		return ExcelUtil.readListFromExcel(file, fieldConverter, fields);
	}
	/**
	 * <xmp>
	 * excel to List<T> 从excel文件读出对象List集合
	 * </xmp>
	 * <br/>1.可自定excel表头字段<br/>2.可自定所需字段<br/>3.可自定字段转换规则
	 * @param file excel文件
	 * @param cls 目标类class
	 * @param fieldConverter 自定转换规则，可读取map的值为新的map，如student.course.name student对象的course为引用类型，可通过定义转换器实现构造出List<Map<studentField,Map<courseField,stringValue>>>
	 * @param tableHeadAlterConfig 对excel表头的列名进行修正，如"日期"修正为date 以匹配实体类字段
	 * @param fields 指定所需的excel表头字段，传入null则转换所有字段
	 * @return List T类型的对象的的List集合
	 * @throws Exception 暂不支持单元格为数字，字符，公式以外的类型
	 */
	public static <T> List<T> readListFromExcel(File file,Class<T> cls,Function<CoverterContext, Object> fieldConverter,Map<Integer,String> tableHeadAlterConfig,String[] fields) throws Exception{
		if(null!=tableHeadAlterConfig){
			alterTableHead(file,tableHeadAlterConfig);
		}
		List<Map<String,Object>> listMap = readListFromExcel(file,fieldConverter,fields);
		List<T> listObj=new ArrayList<>();
		ObjectMapper mapper = new ObjectMapper();
		for (Map<String, Object> map : listMap) {
			listObj.add(mapper.convertValue(map, cls));
		}
		return listObj;
	}
	/**
	 * <xmp>
	 * excel to List<Map<String,Object> 从excel文件读出元素为Map的List集合
	 * </xmp>
	 * <br/>1.可自定所需字段<br/>2.可自定字段转换规则 <br/>3.如需自定excel表头字段可先调用alterTableHead() <br/>3.excel读取以第一列为索引，因此第一列遇到null，或""则结束
	 * @param file excel文件
	 * @param <xmp>fieldConverter 自定义转换规则 Function<String, Object> string为excel表头字段，Object为转换结果</xmp>
	 * @param fields 指定所需的excel表头字段，传入null则转换所有字段
	 * @return K表头字段 V单元格值构成的Map 置入ArrayList
	 * @throws Exception "暂不支持单元格为空白，布尔，错误，数字，字符，公式以外的类型"
	 */
	public static List<Map<String,Object>> readListFromExcel(File file,Function<CoverterContext, Object> fieldConverter,String[] fields) throws Exception{
		Workbook wb = null;
		FormulaEvaluator eval=null;//公式計算器對象
			if (file.getName().endsWith(".xls")) {
				wb = new HSSFWorkbook(new FileInputStream(file));
				eval=new HSSFFormulaEvaluator((HSSFWorkbook) wb);
			} else {
				wb = new XSSFWorkbook(new FileInputStream(file));
				eval=new XSSFFormulaEvaluator((XSSFWorkbook) wb);
			}
			Sheet sheet = wb.getSheetAt(0);
			List list=new ArrayList<HashMap<String,Object>>();
			Row headRow=sheet.getRow(0);
			String[] headFields = new String[MAX_FIELDS_COUNT];
			for (int i = 0; i < headRow.getPhysicalNumberOfCells(); i++) {
				headFields[i]=headRow.getCell(i).getStringCellValue();
			}
			//System.out.println("debug: "+"目標Excel文件物理行數"+sheet.getPhysicalNumberOfRows());
			//遍曆每一行
			for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
				HashMap map = new HashMap<String,Object>();
				Row row=sheet.getRow(i);
				//row==null 空行，如使用了清除内容。遇到空行，或第一列单元格为空，均提前结束
				if(
						null==row
						||row.getCell(0)==null
						||row.getCell(0).getCellType()==Cell.CELL_TYPE_BLANK
						||(row.getCell(0).getCellType()==Cell.CELL_TYPE_STRING&&"".equals(row.getCell(0).getStringCellValue()))
								)break;
				//一行的每一個單元格
				int columnCount=headRow.getPhysicalNumberOfCells()<MAX_FIELDS_COUNT?headRow.getPhysicalNumberOfCells():MAX_FIELDS_COUNT;
				for(int iColumn =0;iColumn<columnCount;iColumn++){
					String cellValue;
					Cell cell= row.getCell(iColumn);
					if(cell!=null){
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_STRING:
							cellValue=row.getCell(iColumn).getStringCellValue();
							break;
						case Cell.CELL_TYPE_NUMERIC:
							//注意日期类型包含在数字类型中，额外判断
							if (HSSFDateUtil.isCellDateFormatted(row.getCell(iColumn))) {
								SimpleDateFormat sdf=new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSSZ");
								cellValue=sdf.format(row.getCell(iColumn).getDateCellValue());
								//System.out.println(cellValue);
								break;
							}
							cellValue=String.valueOf(row.getCell(iColumn).getNumericCellValue());
							break;
						case Cell.CELL_TYPE_FORMULA:
							eval.evaluateFormulaCell(cell);
							cellValue=String.valueOf(row.getCell(iColumn).getNumericCellValue());
							break;
						case Cell.CELL_TYPE_ERROR:
							cellValue=String.valueOf(row.getCell(iColumn).getErrorCellValue());
							break;
						case Cell.CELL_TYPE_BLANK:
							cellValue=null;
							break;
						case Cell.CELL_TYPE_BOOLEAN:
							cellValue=String.valueOf(row.getCell(iColumn).getBooleanCellValue());
							break;
						default:
							throw new Exception("暂不支持单元格为空白，布尔，错误，数字，字符，公式以外的类型");
							//break;
						}
					}else{
						//System.err.println("单元格为null，可将cellvalue置为null");
						cellValue=null;
					}
					//按给定字段提取，并放进map,若自定了转换器，则调用
					if(null!=fields&&fields.length>0){
						List<String> listFeilds=Arrays.asList(fields);
						if(listFeilds!= null&&listFeilds.contains(headFields[iColumn])){
							if(null!=fieldConverter){
								CoverterContext coverterContext = new ExcelUtil.CoverterContext(map,headFields[iColumn],cellValue);
								Object apply = fieldConverter.apply(coverterContext);
								if(coverterContext.getConvertResultType()==ExcelUtil.CoverterContext.CONVERT_RESULT_TYPE_RETURNED){
									map.put(headFields[iColumn],apply);
								}
							}else{
								map.put(headFields[iColumn], cellValue);
							}
						}
					}else{
						/**
						 * 若用户自定了转换规则，则将上下文传递给用户转换函数，并调用
						 */
						if(null!=fieldConverter){
							CoverterContext coverterContext = new ExcelUtil.CoverterContext(map,headFields[iColumn],cellValue);
							Object apply = fieldConverter.apply(coverterContext);
							//转换结果分两种，一种是在转换函数中将结果手动添加到节点上，则无需再处理
							//另一种是仅定义了转化过程，返回了转换结果（默认）
							//带返回结果的
							if(coverterContext.getConvertResultType()==ExcelUtil.CoverterContext.CONVERT_RESULT_TYPE_RETURNED){
								map.put(headFields[iColumn],apply);
							}
						}else{
							map.put(headFields[iColumn], cellValue);
						}
					}
				}
				list.add(map);
			}
			wb.close();
		return list;
	}
	/**
	 * 按照map配置修改表頭,map中key為目標單元格索引,value為值 请注意隐藏列也计数
	 * @param file 文件
	 * @param config 描述excel表头
	 * @throws Exception 文件不存在
	 */
	public static void alterTableHead(File file,Map<Integer, String> config) throws Exception{
		if(null==config||config.size()==0)return;
		Workbook wb = null;
		if (file.getName().endsWith(".xls")) {
			wb = new HSSFWorkbook(new FileInputStream(file));
		} else {
			wb = new XSSFWorkbook(new FileInputStream(file));
		}
		Sheet sheet = wb.getSheetAt(0);
		Row headRow=sheet.getRow(0);
		Set<Integer> keySet=config.keySet();
		for (Integer key : keySet) {
			headRow.getCell(key).setCellValue(config.get(key));
		}
		wb.write(new FileOutputStream(file));
		wb.close();
	}
	/**
	 * 获取对象的指定字段,包括map，list
	 * @param o 源对象
	 * @param field 字段表达式 形如 sss.0||sss.key||name||user.name||user.department.name||sss.key.name
	 * @return 字段对象
	 */
	public static Object getFieldFromObject(Object o,String field) throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException{
		//sss.0.||sss.key||name||user.name||user.department.name||sss.key.name
		//去掉一層 user.name-> name
		if(null==o)return"";
		if("".equals(field))return o;
		String[] fieldGrade=field.split("\\.");
		Object value=null;
		if(o instanceof Map){ 	
			value = ((Map) o).get(fieldGrade[0]);
		}
		if(o instanceof List){ 	
			int i = Integer.parseInt(fieldGrade[0]);
			if(((List) o).size()>i)
			value = ((List) o).get(i);
		}
		if(value==null){ 	
			//反射获取
			try {
				Field fieldRefect = o.getClass().getDeclaredField(fieldGrade[0]);
				fieldRefect.setAccessible(true);
				value = fieldRefect.get(o);
			} catch (Exception e) {
				return null;
			}
		}
		//如果是最後一層，返回
		if(fieldGrade.length-1<=0){
			return value;
		}else{
			return getFieldFromObject(value,field.substring(fieldGrade[0].length()+1));
		}
	}

	/**
	 * 在对象的指定节点上添加字段，若指定节点不存在，则找到指定节点的父节点（要求是map节点），
	 * <br/>添加一个map节点作为指定节点，再在该map上添加字段
	 * <br/>主要用于从excel文件读取数据时，自定义转换规则
	 * @param o 源对象
	 * @param fieldExpLocation 形如sss.0||sss.key||name||user.name||user.department.name||sss.key.name
	 * @param addField 要添加的字段名
	 * @param addition 要添加的字段值
	 * @return 源对象
	 */
	public static Object putFieldToObject(Object o,String fieldExpLocation/*目标节点*/, String addField,Object addition) throws Exception{
		//sss.0.||sss.key||name||user.name||user.department.name||sss.key.name
		int lastIndexOf = fieldExpLocation.lastIndexOf("\\.");
		//目标节点的上一层，即父节点
		String addLocationPerentNodeExp=fieldExpLocation.substring(0,lastIndexOf==-1?0:lastIndexOf+1);
		String addLocationNodeName=fieldExpLocation.substring(lastIndexOf==-1?0:fieldExpLocation.lastIndexOf("\\."));
		
		Object addLocationPerent = ExcelUtil.getFieldFromObject(o,addLocationPerentNodeExp);
		//目标节点不存在，在目标节点的父节点上插入目标节点map
		if (addLocationPerent instanceof Map) {
			if (!((Map) addLocationPerent).containsKey(addLocationNodeName)) {
				((Map) addLocationPerent).put(addLocationNodeName, new HashMap<String, Object>());
			}
		} else {
			throw new Exception("目标节点不存在时，且目标节点的父节点不是Map，无法构建目标节点，并添加字段");
		}
		//添加字段
		Object fieldExpLocationNode = ExcelUtil.getFieldFromObject(o,fieldExpLocation);
		if(fieldExpLocationNode instanceof Map){ 	
			((Map) fieldExpLocationNode).put(addField,addition);
		}else{ 	
			throw new Exception("目前仅支持在map节点上添加,目标节点不是Map");
		}
		return o;
	}
	/************************************导出工具********************************/
	/**
	 * 访问对象的任意字段的string值
	 * 字段表达式如
	 * user.name||user.department.name||list.0.field map.key.field object.field
	 * @param o 源对象
	 * @param field 字段表达式 
	 * @return 对应的字段值
	 */
	public static String getValueFromObject(Object o,String field) throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException{
		//String [] fields={"inTime","isn","inPeople[0].workid","inPeople[0].name","inPeople[0].department","inContact","reason","inCheckPeople[0].name","inCheckPeopleMaterial[0].name","outTime","outPeople[0].workid","outPeople[0].name","outCheckPeople[0].name","outCheckPeopleMaterial[0].name"};
		//sss.0.||sss.key||name||user.name||user.department.name||sss.key.name
		//去掉一層 user.name-> name
		if(null==o)return"";
		String[] fieldGrade=field.split("\\.");
		Object value=null;
		if(o instanceof Map){ 	
			value = ((Map) o).get(fieldGrade[0]);
		}
		if(o instanceof List){ 	
			int i = Integer.parseInt(fieldGrade[0]);
			if(((List) o).size()>i)
			value = ((List) o).get(i);
		}
		if(value==null){ 	
			//反射获取
			try {
				Field fieldRefect = o.getClass().getDeclaredField(fieldGrade[0]);
				fieldRefect.setAccessible(true);
				value = fieldRefect.get(o);
			} catch (Exception e) {
				return null;
			}
		}
		//如果是最後一層，返回
		if(fieldGrade.length-1<=0){
			return null==value?"":value.toString();
		}else{
			return getValueFromObject(value,field.substring(fieldGrade[0].length()+1));
		}
		
	}
	/**
	 * 根據字段表达式數組獲取对象，Map，List的多个值
	 * 字段表达式如
	 * user.name||user.department.name||list.0.field map.key.field object.field
	 * @see ExcelInportUtil.getValueFromObject(Object, String)
	 * @param o 源对象
	 * @param fields 字段数组
	 * @return 字段数组对应的字段值的String表示的List集合
	 */
	public static List<String> getLineFromObject(Object o,String... fields) throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException{
		
			List<String> line=new ArrayList<String>();
				for (String filed : fields) {
					String string = getValueFromObject(o,filed);
					line.add(string);
				}
				return line;
	}
	/**
	 * <xmp>
	 * List<T> to ExcelFile 把T类型的所有字段和值输出excel到指定文件夹,或把T类型指定字段输出Excel到指定文件夹
	 * </xmp>
	 * @param list 源list 仅支持List元素为实体对象
	 * @param outFoderPath 输出文件及
	 * @param fileName 输出文件名，请以xlsx结尾，填入null将使用UUID生成文件名
	 * @param heads 产生的excel文件的表头
	 * @param fields 获取的字段名，可访问复杂字段，如student.courses.2.name,访问学生的第三门课程的名称@see ExcelExportUtil.getValueFromObject(Object, String)
	 * @return Excel文件
	 */
	public static <T> File createExcelFile(List<T> list,String outFoderPath,String fileName,String [] heads, String... fields) throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException {
		File outFile= new File(outFoderPath, fileName==null?(UUID.randomUUID().toString() + ".xlsx"):fileName);
		if(!outFile.getParentFile().exists()){
			outFile.getParentFile().mkdirs();
		}
		Workbook wb;
		OutputStream fos = null;
		
//		int colIndex = 0;
//		Cell cell;
		try {
			wb = new XSSFWorkbook();
			Sheet sheet = wb.createSheet();
			int rowIndex = 0;
			int colIndex=0;
			Row row;
			Cell cell;
			row=sheet.createRow(rowIndex);
			for(String ss:heads){
				cell = row.createCell(colIndex);
				cell.setCellValue(ss);
				colIndex++;
			}
			rowIndex++;
			for (Object o : list) {
				row=sheet.createRow(rowIndex);
				List<String> line=getLineFromObject(o,fields);
				colIndex = 0;
				for(String sss:line){
					cell = row.createCell(colIndex);
					cell.setCellValue(sss);
					colIndex++;
				}
//				line.forEach(rowWrite->{
//					Cell cell = row.createCell(colIndex);
//					cell.setCellValue(rowWrite);
//					colIndex++;
//				});
				rowIndex++;
			}

			fos = new FileOutputStream(outFile);
			wb.write(fos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} finally {
			try {
				if (null != fos)
					fos.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return outFile;
	}
}
