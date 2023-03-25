package excel.excelproject.service;

import excel.excelproject.model.FormModel;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;

public interface MainService {
    public String excelDown(HttpServletRequest request, HttpServletResponse response, FormModel formModel);
}
