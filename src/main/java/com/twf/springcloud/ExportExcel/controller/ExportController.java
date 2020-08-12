package com.twf.springcloud.ExportExcel.controller;

import com.twf.springcloud.ExportExcel.entity.BaseDataResult;
import com.twf.springcloud.ExportExcel.entity.User;
import com.twf.springcloud.ExportExcel.sevice.ExportService;
import com.twf.springcloud.ExportExcel.utils.ExcelPOIUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
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
	@RequestMapping("exportExcel1")
	public ResponseEntity<byte[]> exportExcel1(HttpServletRequest request, HttpServletResponse response) {
		return exportService.exportExcel1(request,response);
	}


	// 导入
	@RequestMapping(value = "importData")
	@ResponseBody
	public BaseDataResult importData(@RequestParam MultipartFile file) {
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
			return new BaseDataResult("400", "导入成功");
		} catch (Exception e) {
			return new BaseDataResult("", e.getMessage());
		}
	}
}
