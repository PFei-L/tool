package tool;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;




import tool.util.commonTool;



public class compareTool {

	public static void main(String[] args) {

		commonTool eu = new commonTool();  
	        
	      //从第几行开始读取  
	      eu.setStartReadPos(0);  
	      File[] fileList_dev=new File("C:\\b\\dev").listFiles();
	      File[] fileList_pro=new File("C:\\b\\pro").listFiles();
	      if(fileList_dev.length!=fileList_pro.length){
	    	  new Exception("File Count Not！");
	      }
	      for(int y=0;y<fileList_dev.length;y++){
	    	  if(!fileList_dev[y].isDirectory()){
	    		//开发环境文件地址
			      File src_xlspath = fileList_dev[y];
			      //生产环境文件地址
			      File src_xlspath_pro =fileList_pro[y];
			      eu.out("FileName："+fileList_dev[y].getName()+"AND"+fileList_pro[y].getName());
			      if(!src_xlspath.getName().equals(src_xlspath_pro.getName())){
			    	  new Exception("Compare Not One");
			      }
			      //差异文件生成地址
			      String dist_xlsPath ="C:\\b\\diffrent"+"\\\\dif_"+fileList_dev[y].getName();
			      ArrayList<String> newList = null;  
			      try {  
			    	  //读取并比较差异
			          newList = eu.readExcel(src_xlspath,src_xlspath_pro); 
			          //生成差异文件
			          eu.writeExcel_xlsx(newList, dist_xlsPath);  
			      } catch (Exception e) {  
			          e.printStackTrace();  
			      }
	    	  }else{
	    		  File[] fileList_dev_n=new File(fileList_dev[y]+"\\").listFiles();
	    		  File[] fileList_pro_n=new File(fileList_pro[y]+"\\").listFiles();
	    	      if(fileList_dev.length!=fileList_pro.length){
	    	    	  new Exception("FileDIR Not ");
	    	      }
	    	      try{
	    	    	  Done(fileList_dev_n,fileList_pro_n,"C:\\b\\diffrent",eu);
	    	      }catch(Exception e){
	    	    	  e.printStackTrace();
	    	      }
	    	      
	    	  }
	      }
	       
	}
	public static void Done(File[] f1,File[] f2,String path,commonTool eu) throws IOException{
		 for(int y=0;y<f1.length;y++){
		   	  if(!f1[y].isDirectory()){
		   		//开发环境文件地址
				      File src_xlspath = f1[y];
				      //生产环境文件地址
				      File src_xlspath_pro =f2[y];
				      eu.out("FileName:"+f1[y].getName()+"AND"+f2[y].getName());
				      if(!src_xlspath.getName().equals(src_xlspath_pro.getName())){
				    	  throw new IOException("Not One File,Please Compare FileCount");
				      }
				      //差异文件生成地址
				      String dist_xlsPath =path+"\\\\dif_"+f1[y].getName();
				      ArrayList<String> newList = null;  
				      try {  
				    	  //读取并比较差异
				          newList = eu.readExcel(src_xlspath,src_xlspath_pro); 
				          //生成差异文件
				          eu.writeExcel_xlsx(newList, dist_xlsPath);  
				      } catch (Exception e) {  
				          e.printStackTrace();  
				      }
		   	  }else{
		   		  File[] fileList_dev_n=new File(f1[y]+"\\").listFiles();
		   		  File[] fileList_pro_n=new File(f2[y]+"\\").listFiles();
		   	      if(f1.length!=f2.length){
		   	    	  throw new IOException("FileCount Not!");
		   	      }
		   	   Done(fileList_dev_n,fileList_pro_n,path,eu);
		   	  }
		     }
	}

}
