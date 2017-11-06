package com.wxstore.commons.excel_utils;

import com.wxstore.annotation.ExcelAttribute;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.nio.channels.FileChannel;
import java.util.List;
import java.util.Map;

/**
 * author:Zan Yang
 * qq:841533078
 */
@Log4j2
public class ExcelUtils {
    /**
     * @param datas 导出报表，以key值作为表名
     * @return
     */
    public static Workbook exportExcelByMaps(Map<String,List<?>> datas){
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        //遍历Map
        for(Map.Entry<String, List<?>> map : datas.entrySet()){
            List<?> list = map.getValue();
            //将key值设为sheet名称
            HSSFSheet sheet = hssfWorkbook.createSheet(map.getKey());
            //判断map的值是否为空
            if(null == map.getValue() || map.getValue().size() == 0)
                continue;
            //使用反射抓取list的泛型
            Type type = list.getClass().getGenericSuperclass();
            ParameterizedType p = (ParameterizedType)type;
            Class cls = (Class)p.getActualTypeArguments()[0];
            //可能有即便list可能传递为空，也需要导出到表格的情况发生，所以弃置定位取类型的方法
            //Field[] fields = list.get(0).getClass().getDeclaredFields();
            Field[] fields = cls.getDeclaredFields();
            Row titleRow = sheet.createRow(0);
            for (int i = 0, m = fields.length; i < m; i++){
                //判断该属性是否是需要导出的列，不是则直接跳过
                if(!fields[i].isAnnotationPresent(ExcelAttribute.class)){
                    continue;
                }
                //获取属性的对应注解
                ExcelAttribute excelAttribute = fields[i].getAnnotation(ExcelAttribute.class);
                if(excelAttribute.isExport()) {
                    int colum = excelAttribute.column() - 1;
                    Cell cell = titleRow.createCell(colum);
                    cell.setCellValue(excelAttribute.name());
                    if(excelAttribute.isAdaptive()) {
                        sheet.autoSizeColumn(colum);
                    }
                }
                continue;
            }
            //遍历list
            int startIndex = 0;
            int endIndex = list.size();
            Row row = null;
            Cell cell = null;
            for(int i = startIndex; i < endIndex; i++){
                row = sheet.createRow(i+1-startIndex);
                //如果list为空或者size为0则跳过
                if(list.size() == 0 || null == list){
                    continue;
                }
                Object o = list.get(i);
                for(int m = 0;m<fields.length;m++){
                    Field field = fields[m];
                    if(!field.isAnnotationPresent(ExcelAttribute.class)){
                        continue;
                    }
                    field.setAccessible(true);
                    ExcelAttribute attr = field.getAnnotation(ExcelAttribute.class);
                    try{
                        //判断该属性在现有需求下是否需要导出
                        if(attr.isExport()){
                            //让列的宽度自适应
                            if(attr.isAdaptive()){
                                sheet.autoSizeColumn(attr.column() - 1);
                            }
                            cell = row.createCell(attr.column() - 1);
                            cell.setCellValue(field.get(o) == null? "":String.valueOf(field.get(o)));
                        }
                    } catch (Exception e){
                        e.printStackTrace();
                    }
                }
            }
        }
        return hssfWorkbook;
    }

    /**
     * 将byte输出至response
     */
    public static void export2Response(Workbook workbook , String fileName,HttpServletResponse response){
        try{
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + new String((fileName + ".xls").getBytes("GBK"), "iso-8859-1"));
            workbook.write(response.getOutputStream());
        } catch (Exception e){
            e.printStackTrace();
        }
    }

    /**
     * 将excel输出至本地路径
     */
    public static void export2Path(Workbook workbook , String fileName,String path){
        File dir = new File(path);
        //路径不存在则创建
        if(!dir.exists()||!(dir.isDirectory())){
            dir.mkdir();
        }
        String filePath = path.concat(fileName).concat(".xls");
        File file = new File(filePath);
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(file);
            workbook.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException ioe){
            ioe.printStackTrace();
        } finally {
            //关闭流
            if(null != out) {
                try {
                    out.flush();
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
