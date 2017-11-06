package com.wxstore.controller;

import com.wxstore.commons.excel_utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Workbook;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class TempController {

    public static Workbook exportExcel(){
        Map<String, List<?>> temp = new HashMap<>();
        UserInfo tempUser1 = new UserInfo("alien","11","1","clone");
        UserInfo tempUser2 = new UserInfo("寿限无寿限无EWQEWQEWQEWQE","2222222","22","dsadsa");
        UserInfo tempUser3 = new UserInfo("dwqdwqdw","222","2","2");
        DogeInfo dogeInfo1 = new DogeInfo("alien","11","1","clone");
        DogeInfo dogeInfo12 = new DogeInfo("寿限无寿限无EWQEWQEWQEWQE","2222222","22","dsadsa");
        DogeInfo dogeInfo13 = new DogeInfo("dwqdwqdw","222","2","2");
        List<UserInfo> users = new ArrayList<UserInfo>(){{add(tempUser1);add(tempUser2);add(tempUser3);}};
        List<DogeInfo> dogeInfos = new ArrayList<DogeInfo>(){{add(dogeInfo13);add(dogeInfo12);add(dogeInfo1);}};
        List<DogeInfo> dogeInfos1 = new ArrayList<>();
        temp.put("user",users);
        temp.put("doge",dogeInfos);
        temp.put("doge2",dogeInfos1);
        temp.put("doge3",null);
        Workbook workbook = ExcelUtils.exportExcelByMaps(temp);
        return  workbook;
    }
    public static void main(String[] args){
        System.out.print(args.toString());
        Workbook workbook = exportExcel();
        ExcelUtils.export2Path(workbook,"temp","D:/File/");
    }
}
