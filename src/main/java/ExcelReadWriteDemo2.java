import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadWriteDemo2 {

    public static void main(String[] args) throws Exception {
        // 设置JVM堆内存大小为2GB
        String inputFileDir = "E:\\excel\\excel8\\"; // 修改为输入文件夹路径
        String outputFile =null;
        String outputFileError =null;
        if(inputFileDir.equals("E:\\excel\\excel1\\")){
            outputFile = "E:\\dataexcel\\Processed_Data1.xlsx"; // 修改为输出文件路径
            outputFileError = "E:\\dataexcel\\outputFileError1.xlsx"; // 修改为输出文件路径
        }
        if(inputFileDir.equals("E:\\excel\\excel2\\")){
            outputFile = "E:\\dataexcel\\Processed_Data2.xlsx"; // 修改为输出文件路径
            outputFileError = "E:\\dataexcel\\outputFileError2.xlsx"; // 修改为输出文件路径

        }
        if(inputFileDir.equals("E:\\excel\\excel3\\")){
            outputFile = "E:\\dataexcel\\Processed_Data3.xlsx"; // 修改为输出文件路径
            outputFileError = "E:\\dataexcel\\outputFileError3.xlsx"; // 修改为输出文件路径

        }
        if(inputFileDir.equals("E:\\excel\\excel4\\")){
            outputFile = "E:\\dataexcel\\Processed_Data4.xlsx"; // 修改为输出文件路径
            outputFileError = "E:\\dataexcel\\outputFileError4.xlsx"; // 修改为输出文件路径

        }
        if(inputFileDir.equals("E:\\excel\\excel5\\")){
            outputFile = "E:\\dataexcel\\Processed_Data5.xlsx"; // 修改为输出文件路径
            outputFileError = "E:\\dataexcel\\outputFileError5.xlsx"; // 修改为输出文件路径

        }
        if(inputFileDir.equals("E:\\excel\\excel6\\")){
            outputFile = "E:\\dataexcel\\Processed_Data6.xlsx"; // 修改为输出文件路径
            outputFileError = "E:\\dataexcel\\outputFileError6.xlsx"; // 修改为输出文件路径

        }
        if(inputFileDir.equals("E:\\excel\\excel7\\")){
            outputFile = "E:\\dataexcel\\Processed_Data7.xlsx"; // 修改为输出文件路径
            outputFileError = "E:\\dataexcel\\outputFileError7.xlsx"; // 修改为输出文件路径

        }
        if(inputFileDir.equals("E:\\excel\\excel8\\")){
            outputFile = "E:\\dataexcel\\Processed_Data8.xlsx"; // 修改为输出文件路径
            outputFileError = "E:\\dataexcel\\outputFileError8.xlsx"; // 修改为输出文件路径

        }

        List<Map<String,String>> list=new ArrayList<>();
        List<String> list2=new ArrayList<>();
        int pp=0;
        File[] files = new File(inputFileDir).listFiles();
        if (files != null) {
            for (File file : files) {
                if (file.isFile() && file.getName().endsWith(".xlsx")) {
                    // 读取输入文件
                    // 调整限制值
                    System.out.println(pp);
                    if(pp==131){
                        System.out.println(pp);
                    }
                    pp++;
                    ZipSecureFile.setMinInflateRatio(0.005);
                    Workbook workbook = WorkbookFactory.create(new FileInputStream(file));
                    Sheet sheet = workbook.getSheetAt(0);
                    // 数据处理
                    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                        Row row = sheet.getRow(i);
                        if (row != null) {
                            if(i==1&&(row.getCell(0)==null||row.getCell(0).toString().equals(""))){
                                list2.add(file.getName());
                                System.out.println("文件不标准： "+file.getName());
                                break;
                            }
                            //存储结果的map
                            if(row.getCell(0)!=null&&!row.getCell(0).toString().equals("")){
                                Map<String,String> resultMap=new HashMap<>();
                                //资源不为null，下一个资源为null，这里面是一个整体
                                resultMap.put("资源名称",row.getCell(0).toString());
                                resultMap.put("会话类型",row.getCell(1).toString());
                                resultMap.put("用户名",row.getCell(2).toString());
                                resultMap.put("账号",row.getCell(3).toString());
                                resultMap.put("服务IP:端口",row.getCell(4).toString());
                                list.add(resultMap);
                                //读取到了之后 ,继续读取下一行
                                continue;
                            }
                            if(row.getCell(1).toString().equals("生成时间")){
                                continue;
                            }
                            try{  //下面是操作读取时间和操作记录
                                if(row.getCell(2)!=null) {
                                    if (row.getCell(2).toString().contains("SET ") ||
                                            row.getCell(2).toString().contains("set ") ||
                                            row.getCell(2).toString().contains("alter ") ||
                                            row.getCell(2).toString().contains("ALTER ") ||
                                            row.getCell(2).toString().contains("UPDATE ") ||
                                            row.getCell(2).toString().contains("update ") ||
                                            row.getCell(2).toString().contains("drop ") ||
                                            row.getCell(2).toString().contains("DROP ") ||
                                            row.getCell(2).toString().contains("vi ") ||
                                            row.getCell(2).toString().contains("mv ") ||
                                            row.getCell(2).toString().contains("vim ")
                                    ) {
                                        Map map = list.get(list.size() - 1);
                                        map.put("操作时间", row.getCell(1).toString());
                                        map.put("操作内容", row.getCell(2).toString());
                                        map.put("need", true);
                                    }
                                }
                            }catch (Exception e){
                                e.printStackTrace();
                                System.out.println(i);
                            }
                        }
                    }
                    workbook.close();
                }
            }
        }

        FileOutputStream outputStream = new FileOutputStream(new File(outputFile));
        FileOutputStream outputStreamError = new FileOutputStream(new File(outputFileError));
        //数据过滤
        // 删除不包含 key 为 "need" 的 Map 数据
        //list=list.stream().filter(e->e.containsKey("need")).collect(Collectors.toList());
        //list = removeMapsWithoutKey(list, "need");
        list.removeIf(map -> !map.containsKey("need"));

        writeExcelFile(outputStream, list);
        writeExcelFile2(outputStreamError, list2);
        System.out.println("数据处理并写入新的 excel 文件完成！");
    }
    public static List<Map<String, String>> removeMapsWithoutKey(List<Map<String, String>> list, String key) {
        List<Map<String, String>> result = new ArrayList<>();
        for (Map<String, String> map : list) {
            if (map.containsKey(key)) {
                result.add(map);
            }
        }
        return result;
    }

    private static void writeExcelFile2(FileOutputStream outputStream, List<String> dataList) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Processed_DataError");
        // 写入表头
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("不标准文件名称");

        // 写入数据
        int rowNum = 1;
        for (String data : dataList) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(data);
        }

        // 自动调整列宽
        for (int i = 0; i < 1; i++) {
            sheet.autoSizeColumn(i);
        }
        System.out.println("开始写入ERROR数据");
        workbook.write(outputStream);
        System.out.println("写入ERROR数据成功");
        workbook.close();
        outputStream.close();
    }


    private static void writeExcelFile(FileOutputStream outputStream, List<Map<String, String>> dataList) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Processed_Data");
        // 写入表头
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("资源名称");
        headerRow.createCell(1).setCellValue("会话类型");
        headerRow.createCell(2).setCellValue("用户名");
        headerRow.createCell(3).setCellValue("账号");
        headerRow.createCell(4).setCellValue("服务IP:端口");
        headerRow.createCell(5).setCellValue("操作时间");
        headerRow.createCell(6).setCellValue("操作内容");

        // 写入数据
        int rowNum = 1;
        for (Map<String, String> data : dataList) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(data.get("资源名称"));
            row.createCell(1).setCellValue(data.get("会话类型"));
            row.createCell(2).setCellValue(data.get("用户名"));
            row.createCell(3).setCellValue(data.get("账号"));
            row.createCell(4).setCellValue(data.get("服务IP:端口"));
            row.createCell(5).setCellValue(data.get("操作时间"));
            row.createCell(6).setCellValue(data.get("操作内容"));
        }

        // 自动调整列宽
        for (int i = 0; i < 7; i++) {
            sheet.autoSizeColumn(i);
        }
        System.out.println("开始写入");
        workbook.write(outputStream);
        System.out.println("写入成功");
        workbook.close();
        outputStream.close();
    }
}
