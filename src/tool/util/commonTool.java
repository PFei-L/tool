package tool.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class commonTool {
	//%%%%%%%%-------常�?部分 开始----------%%%%%%%%%  
	  /** 
	   * 默认的开始读取第一行（索引值为0） 
	   */  
	  private final static int READ_START_POS = 0;  
	    
	  /** 
	   * 默认结�?�读�?�的行�?置为最�?�一行（索引值=0，用负数�?�表示倒数第n行） 
	   */  
	  private final static int READ_END_POS = 0;  
	    
	  /** 
	   * 默认Excel内容的开始比较列�?置为第一列（索引值为0） 
	   */  
	  private final static int COMPARE_POS = 0;  
	    
	  /** 
	   * 默认多文件�?�并的时需�?�?�内容比较（相�?�的内容�?�?�?出现） 
	   */  
	  private final static boolean NEED_COMPARE = true;  
	    
	  /** 
	   * 默认多文件�?�并的新文件�?�到�??称�?�?时，进行覆盖 
	   */  
	  private final static boolean NEED_OVERWRITE = true;  
	    
	  /** 
	   * 默认�?��?作一个sheet 
	   */  
	  private final static boolean ONLY_ONE_SHEET = false;  
	    
	  /** 
	   * 默认读�?�第一个sheet中（�?�有当ONLY_ONE_SHEET = true时有效） 
	   */  
	  private final static int SELECTED_SHEET = 0;  
	    
	  /** 
	   * 默认从第一个sheet开始读�?�（索引值为0） 
	   */  
	  private final static int READ_START_SHEET= 0;  
	    
	  /** 
	   * 默认在最�?�一个sheet结�?�读�?�（索引值=0，用负数�?�表示倒数第n行） 
	   */  
	  private final static int READ_END_SHEET = 0;  
	    
	  /** 
	   * 默认打�?��?��?信�?� 
	   */  
	  private final static boolean PRINT_MSG = true;  
	    
	  //%%%%%%%%-------常�?部分 结�?�----------%%%%%%%%%  
	    

	  //%%%%%%%%-------字段部分 开始----------%%%%%%%%%  
	  /** 
	   * Excel文件路径 
	   */  
	  private String excelPath = "data.xlsx";  

	  /** 
	   * 设定开始读�?�的�?置，默认为0 
	   */  
	  private int startReadPos = READ_START_POS;  

	  /** 
	   * 设定结�?�读�?�的�?置，默认为0，用负数�?�表示倒数第n行 
	   */  
	  private int endReadPos = READ_END_POS;  
	    
	  /** 
	   * 设定开始比较的列�?置，默认为0 
	   */  
	  private int comparePos = COMPARE_POS;  

	  /** 
	   *  设定汇总的文件是�?�需�?替�?�，默认为true 
	   */  
	  private boolean isOverWrite = NEED_OVERWRITE;  
	    
	  /** 
	   *  设定是�?�需�?比较，默认为true(仅当�?覆写目标内容是有效，�?�isOverWrite=false时有效) 
	   */  
	  private boolean isNeedCompare = NEED_COMPARE;  
	    
	  /** 
	   * 设定是�?��?��?作第一个sheet 
	   */  
	  private boolean onlyReadOneSheet = ONLY_ONE_SHEET;  
	    
	  /** 
	   * 设定�?作的sheet在索引值 
	   */  
	  private int selectedSheetIdx =SELECTED_SHEET;  
	    
	  /** 
	   * 设定�?作的sheet的�??称 
	   */  
	  private String selectedSheetName = "";  
	    
	  /** 
	   * 设定开始读�?�的sheet，默认为0 
	   */  
	  private int startSheetIdx = READ_START_SHEET;  

	  /** 
	   * 设定结�?�读�?�的sheet，默认为0，用负数�?�表示倒数第n行     
	   */  
	  private int endSheetIdx = READ_END_SHEET;  
	    
	  /** 
	   * 设定是�?�打�?�消�?� 
	   */  
	  private boolean printMsg = PRINT_MSG;  
	    
	    
	  //%%%%%%%%-------字段部分 结�?�----------%%%%%%%%%  
	    
	  public commonTool(){}  
	  
	    
	  public commonTool(String excelPath){  
	      this.excelPath = excelPath;  
	  }  
	    
	  /** 
	   * 还原设定（其实是�?新new一个新的对象并返回） 
	   * @return 
	   */  
	  public commonTool RestoreSettings(){  
		  commonTool instance = new  commonTool(this.excelPath);  
	      return instance;  
	  }  
	    

	  /** 
	   * 读取并比较文件
	   */  
	  public ArrayList<String> readExcel(File xlsPath,File xlsPath_pro) throws IOException{  
	        
	      try {  
	              return readExcel_xlsx(xlsPath,xlsPath_pro);  
	      } catch (IOException e) {  
	          throw e;  
	      }  
	  }  
	    

	  /** 
	   * 写入Excel（97-03版，xlsx） 
	   */  
	  public void writeDifExcel_xlsx(ArrayList<String> newList, String dist_xlsPath) throws IOException {  

	      // 判断文件路径是否为空  
	      if (dist_xlsPath == null || dist_xlsPath.equals("")) {  
	          out("File Path canot null");  
	          throw new IOException("File Path canot null");  
	      }  

	      if (newList == null || newList.size() == 0) {  
	          out("succsess!");  
	          return;  
	      }  

	          XSSFWorkbook wb = null;  

	          // 判断文件是否存在  
	          File file = new File(dist_xlsPath);  
	          if (file.exists()) {  
	              // 如果文件存在，则删除，重新写入 
	              file.delete();  
	              // 如果文件不存在，则创建一个新的Excel  
	              wb = new XSSFWorkbook();  
	          } else {  
	              // 如果文件不存在，则创建一个新的Excel  
	              wb = new XSSFWorkbook();  
	          }  
	          // 将rowlist的内容写到Excel中  
	          writeExcel(wb, newList, dist_xlsPath);  
	  }  

	  /** 
	   * 写入差异Excel 
	   *  
	   */  
	  public void writeExcel_xlsx(ArrayList<String> newList, String dist_xlsPath) throws IOException {  
	      writeDifExcel_xlsx(newList, dist_xlsPath);  
	  }  


	  /** 
	   * //
	   *  
	   */  
	  public ArrayList<String> readExcel_xlsx(File file,File file_pro) throws IOException {  
	      // 判断文件是否存在  
	      if (!(file.exists() && file_pro.exists())) {  
	          throw new IOException("Excel File not exist！");  
	      }  

	      XSSFWorkbook wb = null;  
	      XSSFWorkbook wb_pro = null; 
	      ArrayList<String> newList = null;  
	      try {  
	          FileInputStream fis = new FileInputStream(file);  
	          FileInputStream fis_pro = new FileInputStream(file_pro);
	          // 去读Excel  
	          wb = new XSSFWorkbook(fis);  
	          wb_pro = new XSSFWorkbook(fis_pro); 

	          newList = readExcel_rel(wb,wb_pro);  

	      } catch (IOException e) {  
	          e.printStackTrace();  
	      }  
	      return newList;  
	  }  



	  /*** 
	   * 读单元格的值 
	   *  
	   */  
	  public static String getCellValue(Cell cell) {  
	      Object result = "";  
	      if (cell != null) {  
	          switch (cell.getCellType()) {  
	          case STRING:
	              result = cell.getStringCellValue(); 
	              break;  
	          case NUMERIC:  
	              result = cell.getNumericCellValue();  
	              break;  
	          case BOOLEAN:  
	              result = cell.getBooleanCellValue();  
	              break;  
	          case FORMULA:  
	              result = cell.getCellFormula();
	              break;  
	          case ERROR:  
	              result = cell.getErrorCellValue();  
	              break;  
	          case BLANK:  
	              break;  
	          default:  
	              break;  
	          }  
	      }  
	      return result.toString();  
	  }  

	  /** 
	   * 通用读Excel 
	 * @throws IOException 
	   *  
	   */  
	  private ArrayList<String> readExcel_rel(Workbook wb,Workbook wb_pro) throws IOException {
		  //新建差异excel对象
	      ArrayList<String> newList=new ArrayList<String>();
	      
	      int sheetCount = 1;//dev文件sheet数 
	      int sheetCount_pro = 1;//生产文件sheet数
	      
	      Sheet sheet = null;  
	      Sheet sheet_pro=null;
	      if(onlyReadOneSheet){  
	          // 获取设定的sheet 
	          sheet =selectedSheetName.equals("")? wb.getSheetAt(selectedSheetIdx):wb.getSheet(selectedSheetName); 
	          sheet_pro =selectedSheetName.equals("")? wb_pro.getSheetAt(selectedSheetIdx):wb_pro.getSheet(selectedSheetName);
	      }else{                          //作多个sheet  
	          sheetCount = wb.getNumberOfSheets();
	          sheetCount_pro = wb_pro.getNumberOfSheets(); 
	      }
	      if(sheetCount!=sheetCount_pro){
	    	  try {
				throw new IOException("sheetCount not！");
			} catch (IOException e) {
				// TODO Auto-generated catch block
			}  
	      }
	      // 处理sheet中的行数据  
	      for(int t=startSheetIdx; t<sheetCount+endSheetIdx;t++){
	    	  String dev=null;
	    	  String pro=null;
	          // 获取设定的sheet  
	          if(!onlyReadOneSheet) {  
	              sheet =wb.getSheetAt(t);  
	              sheet_pro =wb_pro.getSheetAt(t); 
	          }  
	          
	          //获取最大行数
	          int lastRowNum = sheet.getLastRowNum();  
	          int lastRowNum_pro = sheet_pro.getLastRowNum();

	          if(lastRowNum>0){    //如果>0，表示有数据 
	              out("\n start excel"+sheet.getSheetName()+"contant：");  
	          }  
 
              Row row = null; 
	          Row row_pro=null;
	          // 循环读取行dev
	          List<String> devList=cellCompare( row, lastRowNum, sheet, dev);
	          // 循环读取行pro
	          List<String> proList=cellCompare( row_pro, lastRowNum_pro, sheet_pro, pro);
	          //dev行级row集合去重
	          List<String> cutDevList=CutList(devList,proList);
	          //pro行级row集合去重
	          List<String> cutProList=CutList(proList,devList);
	          //处理并合并形成差异文件
	          List<String> dif=function(cutDevList,cutProList);
	          //将差异sheet加入集合
	          if(!("".equals(dif) || dif==null)){
	        	  StringBuffer st=new StringBuffer();
	        	  String str=st.append(sheet.getSheetName()+":"+dif).toString();
	        	  newList.add(str);
	          }
	      }  
		return newList;
	      
	  }  

	  /** 
	   * 写入Excel，保存
	   *  
	   */  
	  private void writeExcel(Workbook wb, ArrayList<String> newList, String dist_xlsPath) {  
		  
	      out("data:"+newList);
	      for(int i=0;i<newList.size();i++){
	    	  if(!("".equals(newList.get(i))||newList.get(i)==null)){
	    		  //解析sheet对象
	    		  String difSheetName=newList.get(i).split(":")[0];;
		    	  String difSheet=newList.get(i).split(":")[1];
	    		  // 创建新的sheet
	              Sheet sheetDifferent =  wb.createSheet(WorkbookUtil.createSafeSheetName(difSheetName));
		    	  //解析row对象
	              String[] difRow=difSheet.split("!,");
		    	  for(int j=0;j<difRow.length;j++){
		    		  if(!("".equals(difRow[j])||difRow[j]==null)){
		    			  String difCell=difRow[j].replace("[", "");
		    			  String Cll=difCell.replace("]", "");
		    			  String Cell=Cll.replace("!", "");
		    			  out(Cell);
		    			  String[] dif=Cell.split(",");
		    			// 创建row, 从第0行开始
                        Row newRwo = sheetDifferent.createRow(j);
                        for(int v=0;v<dif.length;v++){
                        	if(!(dif[v]==null ||"".equals(dif[v]))){
                        		Cell cell=newRwo.createCell(v);
        			            // 写入内容
        		    			cell.setCellValue(dif[v]);
                        	}
                        }
		    		  }
		    	  }
	    	  }
	      }
	   // 写入Excel中 
	      FileOutputStream outputStream = null;
	        
	      try {  
	          outputStream = new FileOutputStream(dist_xlsPath);
	          wb.write(outputStream);
	          outputStream.flush(); 
	      } catch (IOException e) {  
	          out("wirte wrong ");  
	          e.printStackTrace();  
	      }finally{
	    	  try {
				outputStream.close();
				wb.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
	      }  
	  }  


	  /** 
	   * 样式设置
	   *  
	   * @param fromStyle 
	   * @param toStyle 
	   */  
	  public void copyCellStyle(CellStyle fromStyle, CellStyle toStyle) {  
	      toStyle.setAlignment(fromStyle.getAlignment());  
	      // 边框和边框颜色  
	      toStyle.setBorderBottom(fromStyle.getBorderBottom());  
	      toStyle.setBorderLeft(fromStyle.getBorderLeft());  
	      toStyle.setBorderRight(fromStyle.getBorderRight());  
	      toStyle.setBorderTop(fromStyle.getBorderTop());  
	      toStyle.setTopBorderColor(fromStyle.getTopBorderColor());  
	      toStyle.setBottomBorderColor(fromStyle.getBottomBorderColor());  
	      toStyle.setRightBorderColor(fromStyle.getRightBorderColor());  
	      toStyle.setLeftBorderColor(fromStyle.getLeftBorderColor());  

	      // 背景和�?景  
	      toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColor());  
	      toStyle.setFillForegroundColor(fromStyle.getFillForegroundColor());  

	      // 数�?�格�?  
	      toStyle.setDataFormat(fromStyle.getDataFormat());  
	      toStyle.setFillPattern(fromStyle.getFillPattern());  
	      // toStyle.setFont(fromStyle.getFont(null));  
	      toStyle.setHidden(fromStyle.getHidden());  
	      toStyle.setIndention(fromStyle.getIndention());// 首行缩进  
	      toStyle.setLocked(fromStyle.getLocked());  
	      toStyle.setRotation(fromStyle.getRotation());// 旋转  
	      toStyle.setVerticalAlignment(fromStyle.getVerticalAlignment());  
	      toStyle.setWrapText(fromStyle.getWrapText());  

	  }  

	  /** 
	   * 获取合并单元格的值 
	   *  
	   * @param sheet 
	   * @param row 
	   * @param column 
	   * @return 
	   */  
	  public void setMergedRegion(Sheet sheet) {  
	      int sheetMergeCount = sheet.getNumMergedRegions();  

	      for (int i = 0; i < sheetMergeCount; i++) {  
	          // 获取单元格格式设置 
	          CellRangeAddress ca = sheet.getMergedRegion(i);  
	          int firstRow = ca.getFirstRow();  
	          if (startReadPos - 1 > firstRow) {// 如果第一个�?�并�?�元格格�?在正�?数�?�的上�?�，则跳过。  
	              continue;  
	          }  
	          int lastRow = ca.getLastRow();  
	          int mergeRows = lastRow - firstRow;// �?�并的行数  
	          int firstColumn = ca.getFirstColumn();  
	          int lastColumn = ca.getLastColumn();  
	          // 根�?��?�并的�?�元格�?置和大�?，调整所有的数�?�行格�?，  
	          for (int j = lastRow + 1; j <= sheet.getLastRowNum(); j++) {  
	              // 设定�?�并�?�元格  
	              sheet.addMergedRegion(new CellRangeAddress(j, j + mergeRows, firstColumn, lastColumn));  
	              j = j + mergeRows;// 跳过已�?�并的行  
	          }  

	      }  
	  }  
	    

	  /** 
	   * 打�?�消�?�， 
	   * @param msg 消�?�内容 
	   * @param tr �?�行 
	   */  
	  public void out(String msg){  
	      if(printMsg){  
	          out(msg,true);  
	      }  
	  }  
	  /** 
	   * 打�?�消�?�， 
	   * @param msg 消�?�内容 
	   * @param tr �?�行 
	   */  
	  private void out(String msg,boolean tr){  
	      if(printMsg){
	    	  String newStr=null;
	    	  byte[] b =msg.getBytes();
	    	  try {
				newStr=new String(b,"GBK");
			} catch (UnsupportedEncodingException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
	          System.out.print(newStr+(tr?"\n":""));  
	      }  
	  }  
	  /** 
	   * 将单元格元素导出生成行字符串 
	   * @param msg 消�?�内容 
	   * @param tr �?�行 
	   */  
	  private List<String> cellCompare(Row row,int lastRowNum,Sheet sheet,String ro){
		  List<String> fi=new ArrayList<String>();
		  for (int i = startReadPos; i <= lastRowNum + endReadPos; i++) {  
        	  List<String> rowList=null;
              row=sheet.getRow(i);
              if (row != null) {
                   // 进行比较
                   for (int j = 0; j < row.getLastCellNum(); j++) {  
                       String value_pro=getCellValue(row.getCell(j));
                    	   out("for"+(i+1)+"row：not compare！");
                    	   //创建单元格List
                    	   if(rowList==null){
                    		   rowList=new ArrayList<String>();
                    	   }
                    	   //加入差异单元
                    	   rowList.add(value_pro);
                   }
              }
              if(!(rowList==null ||"".equals(rowList))){
            	  String rowStr=rowList.toString();
            	  StringBuffer r=new StringBuffer();
            	  ro=r.append(rowStr+"!").toString();
            	  fi.add(ro);
              }
          }
		return fi;
	  }
	  
	//行级row集合去重
	  private List<String> CutList(List<String> List1 ,List<String> List2 ){
		  LinkedList<String> dif_dev = new LinkedList<>(List1);
	      HashSet<String> set = new HashSet<>(List2);
	      Iterator<String> itor = dif_dev.iterator();
	      while(itor.hasNext()){
	          if(set.contains(itor.next())){
	              itor.remove();
	          }
	      }
		return dif_dev;
	  }
      //区分数据并合并
	  private List<String> function(List<String> list1,List<String> list2){
		  List<String> dif=new ArrayList<>();
		  for(String lt1 :list1){
			  if(!(lt1==null ||"".equals(lt1))){
				  StringBuffer str1=new StringBuffer();
				  String row1=str1.append("dev,"+lt1).toString();
				  dif.add(row1);
			  }
		  }
		  for(String lt2 :list2){
			  if(!(lt2==null||"".equals(lt2))){
				  StringBuffer str2=new StringBuffer();
				  String row2=str2.append("pro,"+lt2).toString();
				  dif.add(row2);
			  }
			  
		  }
		return dif;
	  }
      
	  public String getExcelPath() {  
	      return this.excelPath;  
	  }  

	  public void setExcelPath(String excelPath) {  
	      this.excelPath = excelPath;  
	  }  

	  public boolean isNeedCompare() {  
	      return isNeedCompare;  
	  }  

	  public void setNeedCompare(boolean isNeedCompare) {  
	      this.isNeedCompare = isNeedCompare;  
	  }  

	  public int getComparePos() {  
	      return comparePos;  
	  }  

	  public void setComparePos(int comparePos) {  
	      this.comparePos = comparePos;  
	  }  

	  public int getStartReadPos() {  
	      return startReadPos;  
	  }  

	  public void setStartReadPos(int startReadPos) {  
	      this.startReadPos = startReadPos;  
	  }  

	  public int getEndReadPos() {  
	      return endReadPos;  
	  }  

	  public void setEndReadPos(int endReadPos) {  
	      this.endReadPos = endReadPos;  
	  }  

	  public boolean isOverWrite() {  
	      return isOverWrite;  
	  }  

	  public void setOverWrite(boolean isOverWrite) {  
	      this.isOverWrite = isOverWrite;  
	  }  

	  public boolean isOnlyReadOneSheet() {  
	      return onlyReadOneSheet;  
	  }  

	  public void setOnlyReadOneSheet(boolean onlyReadOneSheet) {  
	      this.onlyReadOneSheet = onlyReadOneSheet;  
	  }  

	  public int getSelectedSheetIdx() {  
	      return selectedSheetIdx;  
	  }  

	  public void setSelectedSheetIdx(int selectedSheetIdx) {  
	      this.selectedSheetIdx = selectedSheetIdx;  
	  }  

	  public String getSelectedSheetName() {  
	      return selectedSheetName;  
	  }  

	  public void setSelectedSheetName(String selectedSheetName) {  
	      this.selectedSheetName = selectedSheetName;  
	  }  

	  public int getStartSheetIdx() {  
	      return startSheetIdx;  
	  }  

	  public void setStartSheetIdx(int startSheetIdx) {  
	      this.startSheetIdx = startSheetIdx;  
	  }  

	  public int getEndSheetIdx() {  
	      return endSheetIdx;  
	  }  

	  public void setEndSheetIdx(int endSheetIdx) {  
	      this.endSheetIdx = endSheetIdx;  
	  }  

	  public boolean isPrintMsg() {  
	      return printMsg;  
	  }  

	  public void setPrintMsg(boolean printMsg) {  
	      this.printMsg = printMsg;  
	  }  
	  
	  
	  
	    
	 
	}  

