package com.twf.springcloud.ExportExcel.controller;

import com.twf.springcloud.ExportExcel.entity.User;
import com.twf.springcloud.ExportExcel.sevice.ExportService;
import com.twf.springcloud.ExportExcel.utils.ExcelPOIUtils;
import com.twf.springcloud.ExportExcel.utils.ExcelWriterUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/exportExcel/")
public class ExportController {
	
	@Autowired
	private ExportService exportService;


	// 导出excel优化方法
	@RequestMapping("exportExcel")
	public ResponseEntity<byte[]> exportExcel(HttpServletRequest request, HttpServletResponse response) {
		return exportService.exportExcel(request,response);
	}


	// 导入
	@RequestMapping(value = "importData")
	@ResponseBody
	public ResponseEntity<byte[]> importData(HttpServletRequest request) {
		InputStream is = null;
		MultipartHttpServletRequest req = (MultipartHttpServletRequest) request;
		MultipartFile file = req.getFile("1.xlsx");
		try {
			//错误maps
			Map<Integer, String> errorMap = new HashMap<>();
			//错误列表
			List<User> errors = new ArrayList<User>();
			//站点数据列表
			List<User> stations = ExcelPOIUtils.readExc(file, User.class,errorMap,errors,6);
			HashMap<String, Object> maps = new HashMap<>();
			maps.put("succ", stations);
			maps.put("fail", errorMap);

			is = new ByteArrayInputStream(maps.toString().getBytes());
			return ExcelWriterUtil.buildResponseEntity(is, "1.xlsx");


		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}
}
