package com.twf.springcloud.ExportExcel.controller;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.twf.springcloud.ExportExcel.sevice.ExportService;

@RestController
@RequestMapping("/exportExcel/")
public class ExportController {
	
	@Autowired
	private ExportService exportService;

	// 导出excel
	@RequestMapping("exportExcel")
	public ResponseEntity<byte[]> exportExcel(HttpServletRequest request, HttpServletResponse response) {
		return exportService.exportExcel(request,response);
	}
}
