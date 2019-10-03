package difference;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Difference {
	public static void main(String[] args) throws Exception {
		
		poiTestMethod(new FileInputStream("target.xlsx"));
	}
	
	public static void poiTestMethod(InputStream ins) throws Exception {
		List<String> c0 = new ArrayList<>(200);
		List<String> c1 = new ArrayList<>(200);
		List<String> c2 = new ArrayList<>(200);
		List<String> c3 = new ArrayList<>(200);
		List<DiffModel> diffList = new ArrayList<>();
		
        //打开XSSFWorkbook
        XSSFWorkbook wb = new XSSFWorkbook(ins);//获取HSSWorkbook
        int sheetsNum = wb.getNumberOfSheets();//excel sheet
         //解析sheet的数据      
        for(int i = 0 ; i < sheetsNum ; i++){
            Map map = new HashMap();
            int columnNum = 0;
            XSSFSheet sh = wb.getSheetAt(i);//得到sheet
            //解析行
            int rowNum = 0;
            rowNum = sh.getLastRowNum() + 1;
            for(int j = 1 ; j <= sh.getLastRowNum() ; j++){
                XSSFRow row = sh.getRow(j);
                if(row == null){
                    rowNum --;
                    continue;
                }
                if("".equals(getStringValue(row, 0)) && "".equals(getStringValue(row, 1)) && "".equals(getStringValue(row, 2)) && "".equals(getStringValue(row, 3))){
                	break;
                }
                //解析列
                c0.add(getStringValue(row, 0));
                c1.add(getStringValue(row, 1));
                c2.add(getStringValue(row, 2));
                c3.add(getStringValue(row, 3));
            }
        }
        int totleCount = c0.size() > c2.size() ? c0.size() : c2.size();
        int c0Index = 0, c2Index = 0;
        for(int i = 0;i < totleCount * 2 + 1;i++){
        	if(c0Index == c0.size()){
        		while(c2Index < c2.size()){
					DiffModel diffModel = new DiffModel("", "", c2.get(c2Index), c3.get(c2Index));
        			diffList.add(diffModel);
        			c2Index++;
				}
        		break;
        	}else if(c2Index == c2.size()){
        		while(c0Index < c0.size()){
        			DiffModel diffModel = new DiffModel(c0.get(c0Index), c1.get(c0Index), "", "");
        			diffList.add(diffModel);
        			c0Index++;
				}
        		break;
        	}
        	if(c0.get(c0Index).equals(c2.get(c2Index))){
        		if(isNumEqual(c1.get(c0Index),c3.get(c2Index))){
        		}else{
        			DiffModel diffModel = new DiffModel(c0.get(c0Index), c1.get(c0Index), c2.get(c2Index), c3.get(c2Index));
        			diffList.add(diffModel);
        		}
        		c0Index++;
        		c2Index++;
        	}else{
        		boolean c2more = false;
        		for(int j = c2Index + 1;j < c2.size();j++){
        			if(c0.get(c0Index).equals(c2.get(j))){
        				if(isNumEqual(c1.get(c0Index),c3.get(j))){
                		}else{
                			DiffModel diffModel = new DiffModel(c0.get(c0Index), c1.get(c0Index), c2.get(j), c3.get(j));
                			diffList.add(diffModel);
                		}
        				while(c2Index < j){
        					DiffModel diffModel = new DiffModel("", "", c2.get(c2Index), c3.get(c2Index));
                			diffList.add(diffModel);
                			c2Index++;
        				}
        				c2more = true;
        				c0Index++;
        				c2Index++;
        				break;
        			}
        		}
        		if(!c2more){
        			DiffModel diffModel = new DiffModel(c0.get(c0Index), c1.get(c0Index), "", "");
        			diffList.add(diffModel);
        			c0Index++;
        		}
        	}
        }
        File file = new File("diff.txt");
        //如果文件不存在，创建一个文件
        if (!file.exists()) {
            file.createNewFile();
        }
        FileWriter fw = null;
        BufferedWriter bw = null;
        try {
            fw = new FileWriter(file);
            StringBuilder sb = new StringBuilder();
            for(DiffModel m : diffList){
            	sb.append(m.toString()).append("\r\n");
            }
            bw = new BufferedWriter(fw);
            bw.write(sb.toString());
        } catch (IOException e) {

        } finally {
            bw.close();
            fw.close();
        }

        System.out.println("success");
        
    }

	private static String getStringValue(XSSFRow row, int i) {
		String result = "";
		XSSFCell cell = row.getCell(i);
		if(cell != null){
			CellType cellType = cell.getCellType();
			if(cellType == CellType.STRING){
				result = cell.getStringCellValue();
			}else if(cellType == CellType.NUMERIC){
				result = String.valueOf(cell.getNumericCellValue());
			}else if(cellType == CellType.BLANK){
//				System.out.println(i+"isBlank");
			}else{
				System.out.println(i+"isException");
			}
			if(i % 2 == 0 && result != null && result.length() > 5){
				if(!result.contains("-")){
					StringBuffer sb = new StringBuffer();
			        String temp = result;
			        sb.append(result.substring(0, 3)).append("-").append(result.substring(3));
			        result = sb.toString();
				}
			}
		}else{
			System.out.println(i+"isBlank");
		}
		return result;
	}
	
	private static boolean isNumEqual(String s1, String s2){
		BigDecimal c1Num = null;
		BigDecimal c3Num = null;
		try {
			c1Num = new BigDecimal(s1);
		} catch (Exception e) {
			c1Num = new BigDecimal("0");
		}
		try {
			c3Num = new BigDecimal(s2);
		} catch (Exception e) {
			c3Num = new BigDecimal("0");
		}
		if(c1Num.compareTo(c3Num) == 0){
			return true;
		}else{
			return false;
		}
	}
	
}
