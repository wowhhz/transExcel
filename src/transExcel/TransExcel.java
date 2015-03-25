package transExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class TransExcel {
	
	Logger logger = Logger.getLogger(TransExcel.class.getName());

	public static void main(String[] args) throws InvalidFormatException, IOException {
		// TODO Auto-generated method stub
		
		int topScresNum = 4;
		if(args!=null && args.length>0){
			String _arg = args[0];
			String[] _args = _arg.split(":");
			if(_args.length==2){
				if(_args[0].equals("addup")){
					topScresNum = Integer.parseInt(_args[1]);
				}
			}
		}
		
		String inFilepath = "infile/",outFilepath="outfile/";
		TransExcel transExcel = new TransExcel();
		File dir = new File(inFilepath);
		if(!dir.exists()){
			throw new FileNotFoundException("检查文件路径及目录("+inFilepath+")是否存在");
		}
		File[] files = dir.listFiles();
		int x=0;
		for (int i = 0; i < files.length; i++) {
			File file = files[i];
			String filename = file.getName();
			if(!filename.endsWith(".xlsx") && !filename.endsWith(".xls")){
				continue;
			}
			x++;
			String newFileNameExcel = filename.substring(0, filename.lastIndexOf("."))
					+"_trans"+filename.substring(filename.lastIndexOf("."));
			String newFileNameHtml = filename.substring(0, filename.lastIndexOf("."))
					+"_trans.html";
			List<Sheet> list = transExcel.parseExcel(file);
			List resultList = transExcel.transInfo(list,topScresNum);
			File outFileExcel = new File(outFilepath+newFileNameExcel);
			File outFileHtml = new File(outFilepath+newFileNameHtml);
			transExcel.writeExcel(outFileExcel,resultList,topScresNum);
			transExcel.writeHtml(outFileHtml,resultList,topScresNum);
		}
		if(x==0){
			System.out.println("Excel文件不存在,无需输出");
		}
		

	}
	
	public List<Sheet> parseExcel(File file) throws InvalidFormatException, IOException{
		logger.info("解析 "+file.getPath());
		FileInputStream fis = null;
		Workbook book = null;
		fis = new FileInputStream(file);
		book = WorkbookFactory.create(fis);
		List<Sheet> list = new ArrayList<Sheet>();
		try{
			int x = 40;
			for (int i = 0; i < x; i++) {
				list.add(book.getSheetAt(i));
			}
		}catch(IllegalArgumentException e){
			
		}
		
		return list;
	}
	
	public List transInfo(List<Sheet> list,int topScresNum) throws IOException{
		List<List> resultList = new ArrayList(); 
		for (int i = 0; i < list.size(); i++) {
			Sheet sheet = list.get(i);
			if(sheet.getLastRowNum()==0){
				continue;
			}
			int rowNum = sheet.getLastRowNum();
			//读取人信息
			List<String> manList = new ArrayList<String>();
			Map<String,Integer> manMap = new HashMap<String,Integer>();
			Map<String,String> scoreStrMap = new HashMap<String,String>();
			Map<String,Integer> scoreMap = new HashMap<String,Integer>();
			Map<String,int[]> numMap = new HashMap<String,int[]>();
			List<Map> maplist = new ArrayList(); 
			
			for (int j = 0; j <= rowNum; j++) {
				Row row = sheet.getRow(j);
				
				String scoreStr = "", nameStr = "";
				String[] names = null;
				boolean doublePrt = row.getLastCellNum()>14;
				if(doublePrt){
					if(j<2)continue;
						scoreStr = row.getCell(6).getStringCellValue();
						
						names = new String[]{
								row.getCell(2).getStringCellValue(),
								row.getCell(3).getStringCellValue(),
								row.getCell(4).getStringCellValue(),
								row.getCell(5).getStringCellValue()
								};
						String[] names1 = {
								row.getCell(7).getStringCellValue(),
								row.getCell(9).getStringCellValue(),
								row.getCell(11).getStringCellValue(),
								row.getCell(13).getStringCellValue()
						};
						double[] score1 = {
								row.getCell(8).getNumericCellValue(),
								row.getCell(10).getNumericCellValue(),
								row.getCell(12).getNumericCellValue(),
								row.getCell(14).getNumericCellValue()
						};
						for (int k = 0; k < score1.length; k++) {
							countInfo(names1[k], (int)score1[k], manMap, scoreStrMap, manList);
						}
						nameStr = row.getCell(2).getStringCellValue();
						String[] scoresStr = scoreStr.split("-");
					}else{
						if(j==0)continue;
						nameStr = row.getCell(2).getStringCellValue();
						scoreStr = row.getCell(3).getStringCellValue();
						String name1 = row.getCell(5).getStringCellValue();
						String name2 = row.getCell(7).getStringCellValue();
						double score1 = row.getCell(6).getNumericCellValue();
						double score2 = row.getCell(8).getNumericCellValue();
						countInfo(name1, (int)score1, manMap, scoreStrMap,manList);
						countInfo(name2, (int)score2, manMap, scoreStrMap,manList);
						names = nameStr.toLowerCase().split("vs");
						
					
				}
				if(nameStr.trim().length()==0)continue;
				
				String[] scoresStr = scoreStr.split("-");
				
				String _str0 = scoresStr[0];
				String _str1 = scoresStr[1];
				if(_str0.length()>1){
					_str0 = _str0.substring(0, 1);
				}
				if(_str1.length()>1){
					_str1 = _str1.substring(0, 1);
				}
				//当前比分
				int num0 = Integer.parseInt(_str0);
				int num1 = Integer.parseInt(_str1);
				for (int k = 0; k < names.length; k++) {
					int[] _nums = new int[3];
					Arrays.fill(_nums, 0);
					if(numMap.containsKey(names[k].trim())){
						//累计胜负
						_nums = numMap.get(names[k].trim());
					}
					if((!doublePrt && k==0) || (doublePrt && k<2)){
						if(num0>num1){
							_nums[0]++;
						}else if(num0<num1){
							_nums[1]++;
						}else if(num0==num1){
							_nums[2]++;
						}
						numMap.put(names[k].trim(), _nums);
					}else{
						if(num1>num0){
							_nums[0]++;
						}else if(num1<num0){
							_nums[1]++;
						}else if(num1==num0){
							_nums[2]++;
						}
						numMap.put(names[k].trim(), _nums);
					}
					
				}
			}

			String[] manSorts = new String[manList.size()];
			Arrays.fill(manSorts, "");
			//System.out.println(scoreStrMap.toString());
			for (int j = 0; j < manList.size(); j++) {
				String name = manList.get(j);
				String[] scores = scoreStrMap.get(name).split(",");
				for (int k = 0; k < scores.length-1; k++) {
					for (int l = k+1; l < scores.length; l++) {
						int score = Integer.parseInt(scores[k]);
						int _score = Integer.parseInt(scores[l]);
						if(score<_score){
							String _tmp = scores[k];
							scores[k] = scores[l];
							scores[l] = _tmp;
						}
					}
				}
				
				int[] topScores = new int[topScresNum];
				Arrays.fill(topScores, 0);
				for (int k = 0; k < scores.length; k++) {
					int _score = Integer.parseInt(scores[k]);
					for (int l = 0; l < topScores.length; l++) {
						if(_score>topScores[l]){
							topScores[l] = _score;
							break;
						}
					}
				}
				//System.out.print(name+"最高分：");
				int total = 0;
				for (int l = 0; l < topScores.length; l++) {
					total+=topScores[l];
					//System.out.print(topScores[l]+",");
				}
				//System.out.println("="+total);
				scoreMap.put(name, total);
				manSorts[j] = name;
			}
			
			for (int j = 0; j < manSorts.length-1; j++) {
				for (int k = j+1; k < manSorts.length; k++) {
					int score = scoreMap.get(manSorts[j]);
					int _score = scoreMap.get(manSorts[k]);
					if(score<_score){
						String _tmp = manSorts[j];
						manSorts[j] = manSorts[k];
						manSorts[k] = _tmp;
					}
				}
			}

			for (int j = 0; j < manSorts.length; j++) {
				String name = manSorts[j];
				int score = scoreMap.get(name);
				int totalnum = manMap.get(name);
				int[] nums = numMap.get(name);
				if(nums==null){
					throw new IOException("对阵列姓名中找不到对应的名字");
				}
				Map map = new HashMap();
				map.put("name", name);
				map.put("4TopScore", score);
				map.put("totalnum", totalnum);
				map.put("win", nums[0]);
				map.put("lose", nums[1]);
				map.put("draw", nums[2]);
				map.put("winning", (nums[0]*100/totalnum)+"%");//(String.format("%.2f", (double)nums[0]/(double)totalnum*100))+"%");
				
				maplist.add(map);
				//System.out.println(manSorts[j]+","+score+","+nums[0]+","+nums[1]);
			}
			resultList.add(maplist);
		}
		return resultList;
	}
	
	public void writeExcel(File outFile,List<List<Map>> resultList,int topScresNum) throws InvalidFormatException, IOException{
		logger.info("生成 "+outFile.getPath());
		if(!outFile.exists())outFile.createNewFile();
		FileOutputStream fos = new FileOutputStream(outFile);
		XSSFWorkbook book = new XSSFWorkbook();
		
		Font font = null;  
		CellStyle style = null;  
		
		font = book.createFont();  
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);  
		style = book.createCellStyle();  
		style.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);  
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);  
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFont(font);  
		
		CellStyle style1 = book.createCellStyle();  
		style1.setAlignment(CellStyle.ALIGN_CENTER);
		
		for (int i = 0; i < resultList.size(); i++) {
			List<Map> maplist = resultList.get(i);
			Sheet sheet = book.createSheet();
			sheet.setDefaultColumnWidth(17);
			
			Row titlerow = sheet.createRow(0);
			String[] titles = {"排名","姓名","最高积分"+topScresNum+"场累计","出场次数","胜","负","胜率"};
			String[] contents = {"","name","4TopScore","totalnum","win","lose","winning"};
			for (int j = 0; j < titles.length; j++) {
				Cell title = titlerow.createCell(j);
				title.setCellValue(titles[j]);
				title.setCellStyle(style);
			}
			
			for (int j = 0; j < maplist.size(); j++) {
				Map outmap = maplist.get(j);
				Row row = sheet.createRow(j+1);
				for (int l = 0; l < contents.length; l++) {
					Cell cell = row.createCell(l);
					cell.setCellStyle(style1);
					if(l==0){
						cell.setCellValue((Integer)(j+1));
					}else if(l==1 || l==6){
						cell.setCellValue((String)outmap.get(contents[l]));
					}else{
						cell.setCellValue((Integer)outmap.get(contents[l]));
					}
					
					
				}
			}
		}
		book.write(fos);
		fos.close();
		
		
	}
	
	public void writeHtml(File outFile,List<List<Map>> resultList,int topScresNum) throws InvalidFormatException, IOException{
		logger.info("生成 "+outFile.getPath());
		if(!outFile.exists())outFile.createNewFile();
		FileOutputStream fos = new FileOutputStream(outFile);
		
		StringBuffer buffer = new StringBuffer();
		buffer.append("<html><body>\n");
		for (int i = 0; i < resultList.size(); i++) {
			List<Map> maplist = resultList.get(i);
			buffer.append("\t<table class=\"hovertable\">\n");
			
			String[] titles = {"排名","姓名","最高积分"+topScresNum+"场累计","出场次数","胜","负","胜率"};
			String[] contents = {"","name","4TopScore","totalnum","win","lose","winning"};
			buffer.append("\t\t<tr>\n");
			for (int j = 0; j < titles.length; j++) {
				buffer.append("\t\t\t<th>"+titles[j]+"</th>\n");
			}
			buffer.append("\t\t</tr>\n");
			
			for (int j = 0; j < maplist.size(); j++) {
				Map outmap = maplist.get(j);
				buffer.append("\t\t<tr>\n");
				for (int l = 0; l < contents.length; l++) {
					if(l==0){
						buffer.append("\t\t\t<td>"+(j+1)+"</td>\n");
					}else{
						buffer.append("\t\t\t<td>"+outmap.get(contents[l])+"</td>\n");
					}
				}
				buffer.append("\t\t</tr>\n");
			}
			
			buffer.append("\t<table>\n");
		}
		buffer.append("</body></html>");
		fos.write(buffer.toString().getBytes("GBK"));
		fos.close();
		
		
	}
	
	public void countInfo(String name,int score,Map<String,Integer> manMap,Map<String,String> scoreStrMap,List<String> manList){
		if(name.trim().length()==0){
			return;
		}
		int prtNum = 0;
		StringBuffer scoreStr = new StringBuffer();
		if(manMap.containsKey(name)){
			prtNum = manMap.get(name);
		}else{
			manList.add(name);
		}
		if(scoreStrMap.containsKey(name)){
			scoreStr.append(scoreStrMap.get(name));
		}
		if(scoreStr.toString().length()>0){
			scoreStr.append(",");
		}
		manMap.put(name, prtNum+1);
		scoreStr.append(score);
		scoreStrMap.put(name, scoreStr.toString());
	}
	

}
