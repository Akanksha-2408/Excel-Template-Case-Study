package com.ApachePOI.ExcelTemplate.Controller;

import com.ApachePOI.ExcelTemplate.Entity.Employee;
import com.ApachePOI.ExcelTemplate.Service.ServiceClass;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import java.io.IOException;
import java.util.List;

@RestController
public class ControllerClass {

    @Autowired
    ServiceClass service;

    @GetMapping("/download")
    public void downloadFile(HttpServletResponse response) throws IOException {
        service.downloadFile(response);
    }

    @PostMapping(value="/upload", consumes="multipart/form-data")
    public ResponseEntity<String> uploadData(@RequestParam("file") MultipartFile file) throws IOException {
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException("Uploaded file is empty or missing!");
        }
        service.uploadFile(file.getInputStream());
        return ResponseEntity.ok("Excel file data uploaded into database");
    }

    @GetMapping("/read-data")
    public ResponseEntity<List<Employee>> findAll(){
        return ResponseEntity.ok(service.findAll());
    }
}
