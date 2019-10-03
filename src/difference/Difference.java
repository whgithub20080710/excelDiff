package difference;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

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
		// 解析Excel
		resolveExcel(ins, c0, c1, c2, c3);
		// 对比差异
		diff(c0, c1, c2, c3, diffList);
		// 写出差分结果
		writeResult(diffList);
		System.out.println("success");

	}

	private static void diff(List<String> c0, List<String> c1, List<String> c2, List<String> c3,
			List<DiffModel> diffList) {
		int totleCount = c0.size() > c2.size() ? c0.size() : c2.size();
		int c0Index = 0, c2Index = 0;
		for (int i = 0; i < totleCount * 2 + 1; i++) {
			if (c0Index >= c0.size()) {
				while (c2Index < c2.size()) {
					DiffModel diffModel = new DiffModel("", "", c2.get(c2Index), c3.get(c2Index));
					diffList.add(diffModel);
					c2Index++;
				}
				break;
			} else if (c2Index >= c2.size()) {
				while (c0Index < c0.size()) {
					DiffModel diffModel = new DiffModel(c0.get(c0Index), c1.get(c0Index), "", "");
					diffList.add(diffModel);
					c0Index++;
				}
				break;
			}
			if (c0.get(c0Index).equals(c2.get(c2Index))) {
				if (isNumEqual(c1.get(c0Index), c3.get(c2Index))) {
				} else {
					DiffModel diffModel = new DiffModel(c0.get(c0Index), c1.get(c0Index), c2.get(c2Index),
							c3.get(c2Index));
					diffList.add(diffModel);
				}
				c0Index++;
				c2Index++;
			} else {
				boolean c2more = false;
				for (int j = c2Index + 1; j < c2.size(); j++) {
					if (c0.get(c0Index).equals(c2.get(j))) {
						if (isNumEqual(c1.get(c0Index), c3.get(j))) {
						} else {
							DiffModel diffModel = new DiffModel(c0.get(c0Index), c1.get(c0Index), c2.get(j), c3.get(j));
							diffList.add(diffModel);
						}
						while (c2Index < j) {
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
				if (!c2more) {
					DiffModel diffModel = new DiffModel(c0.get(c0Index), c1.get(c0Index), "", "");
					diffList.add(diffModel);
					c0Index++;
				}
			}
		}
	}

	private static void resolveExcel(InputStream ins, List<String> c0, List<String> c1, List<String> c2,
			List<String> c3) throws IOException {
		XSSFWorkbook wb = new XSSFWorkbook(ins);
		XSSFSheet sh = wb.getSheetAt(0);

//		System.out.println("lastRowNum:" + sh.getLastRowNum());
		for (int j = 1; j <= 5000; j++) {
			XSSFRow row = sh.getRow(j);
			if (row == null) {
				break;
			}
			if ("".equals(getStringValue(row, 0)) && "".equals(getStringValue(row, 1))
					&& "".equals(getStringValue(row, 2)) && "".equals(getStringValue(row, 3))) {
				break;
			}
			c0.add(getStringValue(row, 0));
			c1.add(getStringValue(row, 1));
			c2.add(getStringValue(row, 2));
			c3.add(getStringValue(row, 3));
		}
		wb.close();
		ins.close();
	}

	private static void writeResult(List<DiffModel> diffList) throws IOException {
		File file = new File("diff.txt");
		if (!file.exists()) {
			file.createNewFile();
		}
		FileWriter fw = null;
		BufferedWriter bw = null;
		try {
			fw = new FileWriter(file);
			StringBuilder sb = new StringBuilder();
			for (DiffModel m : diffList) {
				sb.append(m.toString()).append("\r\n");
			}
			bw = new BufferedWriter(fw);
			bw.write(sb.toString());
		} catch (IOException e) {
			System.out.println("生成差分结果异常");
			e.printStackTrace();
		} finally {
			bw.close();
			fw.close();
		}
	}

	private static String getStringValue(XSSFRow row, int i) {
		String result = "";
		XSSFCell cell = row.getCell(i);
		if (cell != null) {
			CellType cellType = cell.getCellType();
			if (cellType == CellType.STRING) {
				result = cell.getStringCellValue();
			} else if (cellType == CellType.NUMERIC) {
				result = String.valueOf(cell.getNumericCellValue());
			} else if (cellType == CellType.BLANK) {
			} else {
				System.out.println(i + "isException");
			}
			if (i % 2 == 0 && result != null && result.length() > 5 && !result.contains("-")) {
				StringBuilder sb = new StringBuilder();
				sb.append(result.substring(0, 3)).append("-").append(result.substring(3));
				result = sb.toString();
			}
//			System.out.println(row.getRowNum()+"-"+i + result);
		} else {
			System.out.println(row.getRowNum()+" "+i + "isBlank");
		}
		return result;
	}

	private static boolean isNumEqual(String s1, String s2) {
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
		return c1Num.compareTo(c3Num) == 0 ? true : false;
	}

}
