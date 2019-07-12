package com.test.easyexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;

import org.junit.Test;

public class MyExportTest01 {

	/**
	 * 测试使用easypoi根据模板导出excel
	 * 
	 * 
	 * 包括的内容：
	 * 
	 * 1、生成列表数据
	 * 
	 * 2、填充指定单元格数据
	 * 
	 * 3、设置指定复选框选中
	 * 
	 */
	@Test
	public void testExportExcel01() throws Exception {
		Workbook workbook = null;
		FileOutputStream fos = null;
		try {
			// 获取导出的excel模板
			TemplateExportParams params = new TemplateExportParams("my_template.xlsx");

			// 设置需要导出的数据
			Map<String, Object> excelData = saToExcelData();
			setListData(excelData);

			workbook = ExcelExportUtil.exportExcel(params, excelData);

			// 设置复选框选中/,不选中
			setCheckboxValue(workbook);

			// 导出excel文件到指定目录
			File savefile = new File("D:\\testDir\\0_excel\\01_excel_symbol");
			if (!savefile.exists()) {
				savefile.mkdirs();
			}
			String fileName = "D:\\testDir\\0_excel\\01_excel_symbol\\test001.xlsx";
			fos = new FileOutputStream(fileName);
			workbook.write(fos);

			System.out.println("操作成功");
			System.out.println("生成文件保存目录：" + savefile);
		} catch (Exception e) {
			System.out.println("操作失败");
			e.printStackTrace();
		} finally {
			fos.close();
			workbook.close();
		}
	}

	/**
	 * 获取List数据
	 * 
	 * @param map
	 */
	private void setListData(Map<String, Object> map) {
		List<Map<String, String>> historyData = new ArrayList<Map<String, String>>();
		for (int i = 0; i < 4; i++) {
			Map<String, String> lm = new HashMap<String, String>();
			lm.put("name", "大白兔");
			lm.put("age", "25");
			lm.put("gender", "boy");
			historyData.add(lm);
		}
		map.put("maplist", historyData);
	}

	/**
	 * 获取其他需要显示的数据
	 * 
	 * @return
	 */
	private Map<String, Object> saToExcelData() {
		Map<String, Object> map = new HashMap<String, Object>();
		map.put("school", "北京大学");
		map.put("phone", "1801062647*");
		return map;
	}

	// 使用原生方法设置复选框是否选中
	private void setCheckboxValue(Workbook workbook) {
		Sheet sheet = workbook.getSheetAt(0);

		// 链接的值true/false显示所在的行，下标从0开始
		Row row = sheet.getRow(16);

		// 设置复选框"未婚"不选中，列坐标从0开始
		Cell checkbokCell1 = row.getCell(1);
		checkbokCell1.setCellValue(false);

		// 设置复选框"已婚"选中
		Cell checkbokCell2 = row.getCell(2);
		checkbokCell2.setCellValue(true);

		// 隐藏设置复选框选中行
		row.setZeroHeight(true);
	}

}
