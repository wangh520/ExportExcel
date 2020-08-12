package com.twf.springcloud.ExportExcel.sevice.impl;

import com.twf.springcloud.ExportExcel.po.User;
import com.twf.springcloud.ExportExcel.sevice.ExportService;
import com.twf.springcloud.ExportExcel.utils.ExcelWriterUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Service
public class ExportServiceImpl implements ExportService{
	Logger logger = LoggerFactory.getLogger(ExportServiceImpl.class);

	@Override
	public ResponseEntity<byte[]> exportExcel1(HttpServletRequest request, HttpServletResponse response) {
		try {
			logger.info(">>>>>>>>>>开始导出excel>>>>>>>>>>");
			// 造几条数据
			List<User> list = new ArrayList<User>();
			list.add(new User("唐三藏", "男", 30, "13411111111", "东土大唐", "取西经"));
			list.add(new User("孙悟空", "男", 29, "13411111112", "菩提院", "打妖怪"));
			list.add(new User("猪八戒", "男", 28, "13411111113", "高老庄", "偷懒"));
			list.add(new User("沙悟净", "男", 27, "13411111114", "流沙河", "挑担子"));

			// 每一列字段名
			String[] title = new String[] {"序号", "姓名", "性别", "年龄", "手机号", "地址","爱好" };

			// 字段名所在表格的宽度
			int[] titleLength = new int[] {5000, 5000, 5000, 5000, 5000, 5000, 5000 };

			String tableName = "用户表.xlsx";

			Workbook workbook = new SXSSFWorkbook();
			// 如需生成xls的Excel，请使用下面的工作簿对象，注意后续输出时文件后缀名也需更改为xls
			//Workbook workbook = new HSSFWorkbook();

			// 写入数据到工作簿对象内
			InputStream workbookInputStream = exportData1(workbook,list, tableName, title, titleLength);

			return ExcelWriterUtil.buildResponseEntity(workbookInputStream, tableName);
		} catch (Exception e) {
			e.printStackTrace();
			logger.error(">>>>>>>>>>导出excel 异常，原因为：" + e.getMessage());
		}
		return null;
	}
	/**
	 * 生成Excel并写入数据信息
	 * @param dataList 数据列表
	 * @return 写入数据后的工作簿对象
	 */
	public static InputStream exportData1(Workbook workbook,List<User> dataList,String tableName,String[] title,int[] titleLength){
		ByteArrayOutputStream output = null;
		InputStream inputStream = null;
		// 生成xlsx的Excel

		// 生成Sheet表，写入第一行的列头
		Sheet sheet = ExcelWriterUtil.buildDataSheet(workbook,title,titleLength,tableName);

		CellStyle content = ExcelWriterUtil.contentStyle(workbook);// 报表体样式
		//构建每行的数据内容
		int rowNum = 2;//从第三行开始写数据
		for (Iterator<User> it = dataList.iterator(); it.hasNext(); ) {
			User data = it.next();
			if (data == null) {
				continue;
			}
			//输出行数据
			Row row = sheet.createRow(rowNum);
			convertDataToRow(data, row,rowNum,content);
			rowNum++;
		}

		try {
			output = new ByteArrayOutputStream();
			workbook.write(output);
			inputStream = new ByteArrayInputStream(output.toByteArray());
			output.flush();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {

			try {
				if (output != null) {
					output.close();
					if (inputStream != null)
						inputStream.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}

		}
		return inputStream;
	}
	/**
	 * 将数据转换成行
	 * @param user 源数据
	 * @param row 行对象
	 * @return
	 */
	private static void convertDataToRow(User user, Row row,int rowNum,CellStyle content){

		int cellNum = 0;
		Cell cell;

		cell = row.createCell(cellNum++);
		cell.setCellValue(rowNum-1); // 序号
		cell.setCellStyle(content);

		cell = row.createCell(cellNum++);
		cell.setCellValue(user.getName()); // 姓名
		cell.setCellStyle(content);

		cell = row.createCell(cellNum++);
		cell.setCellValue(user.getSex()); // 性别
		cell.setCellStyle(content);

		cell = row.createCell(cellNum++);
		cell.setCellValue(user.getAge()); // 年龄
		cell.setCellStyle(content);

		cell = row.createCell(cellNum++);
		cell.setCellValue(user.getPhoneNo()); // 手机号
		cell.setCellStyle(content);

		cell = row.createCell(cellNum++);
		cell.setCellValue(user.getAddress()); // 地址
		cell.setCellStyle(content);

		cell = row.createCell(cellNum++);
		cell.setCellValue(user.getHobby()); // 爱好
		cell.setCellStyle(content);
	}

}
