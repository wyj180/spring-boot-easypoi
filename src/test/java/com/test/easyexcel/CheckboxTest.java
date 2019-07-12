package com.test.easyexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;

import org.junit.Test;

public class CheckboxTest {

	/**
	 * 测试根据模板导出excel时设置复选框选中或不选中
	 * 
	 */
	@Test
	public void testCheckbox() throws Exception {
		Workbook workbook = null;
		FileOutputStream fos = null;
		try {
			// 获取导出的excel模板
			TemplateExportParams params = new TemplateExportParams("checkbox_template.xlsx");

			// 设置需要导出的数据
			Map<String, Object> excelData = new HashMap<>();

			workbook = ExcelExportUtil.exportExcel(params, excelData);

			// 设置复选框选中/,不选中
			setCheckboxValue(workbook);

			// 导出excel文件到指定目录
			File savefile = new File("D:\\testDir\\0_excel\\01_excel_symbol");
			if (!savefile.exists()) {
				savefile.mkdirs();
			}
			String fileName = "D:\\testDir\\0_excel\\01_excel_symbol\\checkbox_test001.xlsx";
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
	private void setCheckboxValue(Workbook workbook) {
		Sheet sheet = workbook.getSheetAt(0);

		// 链接的值true/false显示所在的行，下标从0开始
		Row row = sheet.getRow(2);

		// 设置复选选中
		Cell checkbokCell1 = row.getCell(5);
		checkbokCell1.setCellValue(true);
	}

}
