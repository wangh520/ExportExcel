package com.twf.springcloud.ExportExcel.controller;

import com.twf.springcloud.ExportExcel.sevice.ExportService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

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
}
