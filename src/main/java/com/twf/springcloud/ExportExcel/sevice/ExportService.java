package com.twf.springcloud.ExportExcel.sevice;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.http.ResponseEntity;

public interface ExportService {

	ResponseEntity<byte[]> exportExcel(HttpServletRequest request, HttpServletResponse response);

}
