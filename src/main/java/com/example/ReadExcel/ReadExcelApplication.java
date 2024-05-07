package com.example.ReadExcel;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;


@SpringBootApplication
public class ReadExcelApplication {

    public static void main(String[] args) throws IOException {
        SpringApplication.run(ReadExcelApplication.class, args);

        Service sv = new Service();
        sv.ReadTable();

    }


}
