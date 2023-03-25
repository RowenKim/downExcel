package excel.excelproject.service;

import excel.excelproject.model.FormModel;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.net.URLEncoder;

@Service
public class MainServiceImpl implements MainService{


    @Override
    public String excelDown(HttpServletRequest request, HttpServletResponse response, FormModel formModel) {
        try {
            Workbook workbook = new HSSFWorkbook();
            Sheet sheet = workbook.createSheet();

            // 파일명 설정
            String tempName = formModel.getFileName() + ".xls";
            int rowNum = formModel.getRowNum();
            int colNum = formModel.getColNum();

            String name = "출력될파명이름";
            // 한글이 포함

            if(tempName.matches(".*[ㄱ-ㅎㅏ-ㅣ가-힣]+.*")) {
                // 브라우저 별 한글 인코딩 https://byul91oh.tistory.com/487 참고
                String header = request.getHeader("User-Agent");
                if (header.contains("Edge")){
                    name = URLEncoder.encode(tempName, "UTF-8").replaceAll("\\+", "%20");
                    response.setHeader("Content-Disposition", "attachment;filename=\"" + name + "\".xlsx;");
                } else if (header.contains("MSIE") || header.contains("Trident")) { // IE 11버전부터 Trident로 변경되었기때문에 추가해준다.
                    name = URLEncoder.encode(tempName, "UTF-8").replaceAll("\\+", "%20");
                    response.setHeader("Content-Disposition", "attachment;filename=" + name + ".xlsx;");
                } else if (header.contains("Chrome")) {
                    name = new String(tempName.getBytes("UTF-8"), "ISO-8859-1");
                    response.setHeader("Content-Disposition", "attachment; filename=\"" + name + "\".xlsx");
                } else if (header.contains("Opera")) {
                    name = new String(tempName.getBytes("UTF-8"), "ISO-8859-1");
                    response.setHeader("Content-Disposition", "attachment; filename=\"" + name + "\".xlsx");
                } else if (header.contains("Firefox")) {
                    name = new String(tempName.getBytes("UTF-8"), "ISO-8859-1");
                    response.setHeader("Content-Disposition", "attachment; filename=" + name + ".xlsx");
                }
            } else {// 한글 미 포함
                name = tempName;
            }

            // 입력 Row 만큼 반복문 실행
            for (int rowCnt = 0; rowCnt < rowNum; rowCnt++){
                Row row = sheet.createRow(rowCnt);
                // 입력 Column 만큼 반복문 실행
                for(int colCnt = 0; colCnt < colNum; colCnt++){
                    row.createCell(colCnt).setCellValue("cell Number : " + colCnt);
                }
            }

            response.setContentType("ms-vnd/excel");
            response.setHeader("Content-Disposition", "attachment;filename=" + name);

            workbook.write(response.getOutputStream());
            workbook.close();

        } catch (IOException e) {
            throw new RuntimeException(e);
        }


        return null;
    }
}
