package excel.excelproject.controller;

import excel.excelproject.model.FormModel;
import excel.excelproject.service.MainServiceImpl;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import java.util.HashMap;
import java.util.Map;

@Controller
public class MainController {

    @Autowired
    MainServiceImpl mainService;

    @GetMapping("/")
    public String mainPage(){
        return "index.html";
    }

    @PostMapping("/excelDown")
    public @ResponseBody Map<Object, Object> excelDown(HttpServletRequest request, HttpServletResponse response, FormModel formModel){
        HashMap<Object, Object> map = new HashMap<>();
        String result = mainService.excelDown(request, response, formModel);
        map.put("result", result);
        return  map;
    };
}
