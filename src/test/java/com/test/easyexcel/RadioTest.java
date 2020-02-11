package com.test.easyexcel;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

public class RadioTest {

	/**
	 * 根据模板导出excel时设置单元框选中
	 * 
	 */
	@Test
	public void testRadio() throws Exception {
		Workbook workbook = null;
		FileOutputStream fos = null;
		try {
			// 获取导出的excel模板
			TemplateExportParams params = new TemplateExportParams("radio_template.xlsx");

			// 设置需要导出的数据
			Map<String, Object> excelData = new HashMap<>();

			workbook = ExcelExportUtil.exportExcel(params, excelData);

			// 设置复选框选中/,不选中
			setRadioValue(workbook);

			// 导出excel文件到指定目录
			File savefile = new File("D:\\testDir\\0_excel\\01_excel_symbol");
			if (!savefile.exists()) {
				savefile.mkdirs();
			}
			String fileName = "D:\\testDir\\0_excel\\01_excel_symbol\\radio_test001.xlsx";
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

	// 使用原生方法设置复选框是否选中
	private void setRadioValue(Workbook workbook) {
		Sheet sheet = workbook.getSheetAt(0);

		// 获取单选框所在行的下标，下标从0开始
		Row row = sheet.getRow(12);

		// 设置单元框的值，1表示选中第一个单选框，2表示选中第二个单选框，下标从0开始
		Cell radioCell1 = row.getCell(6);
		radioCell1.setCellValue(3);
	}

}
