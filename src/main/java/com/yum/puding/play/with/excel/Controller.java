package com.yum.puding.play.with.excel;

import java.io.*;
import java.util.Base64;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.core.io.ByteArrayResource;



@RestController
@RequestMapping("/xlsx")
public class Controller {
    private static final Logger logger = LogManager.getLogger(Controller.class);

    @PostMapping("/decode")
    public ResponseEntity<Object> base64ToExcel(@RequestBody Base64Req request){
        byte[] fileBytes = Base64.getDecoder().decode(request.getFile());

        try(
            InputStream inputStream = new ByteArrayInputStream(fileBytes);
            Workbook workbook = WorkbookFactory.create(inputStream);
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream()
        ) {
            workbook.write(outputStream);
            Resource resource = new ByteArrayResource(outputStream.toByteArray());
            return ResponseEntity.ok().contentType(MediaType.parseMediaType("application/octet-stream")).header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=test.xlsx;").body(resource);
        }catch (Exception e){
            e.printStackTrace();
            return new ResponseEntity<>(HttpStatus.OK);
        }

    }

    @PostMapping("/encode")
    public ResponseEntity<Object> excelToBase64(@RequestParam("file") MultipartFile file){

        try(
                Workbook workbook = WorkbookFactory.create(file.getInputStream());
                ByteArrayOutputStream outputStream = new ByteArrayOutputStream()
        ) {
            workbook.write(outputStream);
            byte[] fileBytes = outputStream.toByteArray();
            String fileBase64 = Base64.getEncoder().encodeToString(fileBytes);

            Base64Req resData = new Base64Req();
            resData.setFile(fileBase64);

            return new ResponseEntity<>(resData, HttpStatus.OK);
        }catch (Exception e){
            e.printStackTrace();
            return new ResponseEntity<>(HttpStatus.BAD_REQUEST);
        }

    }
}
