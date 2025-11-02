package com.ApachePOI.ExcelTemplate.Controller;

import com.ApachePOI.ExcelTemplate.Entity.Employee;
import com.ApachePOI.ExcelTemplate.Service.ServiceClass;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
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

    @PostMapping("/upload")
    public ResponseEntity<?> uploadData(@RequestParam("file") MultipartFile file) { // Removed HttpServletResponse injection
        try {
            byte[] errorFileBytes = service.processUpload(file);

            if (errorFileBytes == null) {
                return ResponseEntity.ok("File uploaded and data is valid. Employees saved successfully.");
            } else {
                HttpHeaders headers = new HttpHeaders();
                headers.setContentDispositionFormData("attachment", "Error_Template.xlsx");
                headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

                return new ResponseEntity<>(errorFileBytes, headers, HttpStatus.BAD_REQUEST);
            }
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("An error occurred during file processing: " + e.getMessage());
        }
    }

    @GetMapping("/read-data")
    public ResponseEntity<List<Employee>> findAll(){
        return ResponseEntity.ok(service.findAll());
    }
}
