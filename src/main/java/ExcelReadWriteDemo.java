import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadWriteDemo {

    public static void main(String[] args) throws Exception {
        // 设置JVM堆内存大小为2GB
        String inputFile1 = "C:\\Users\\fh\\Documents\\WXWork\\1688850891728517\\Cache\\File\\2023-04\\";
        String inputFile = inputFile1+"new.xlsx";

        // 读取输入文件
        Workbook workbook = WorkbookFactory.create(new FileInputStream(new File(inputFile)));
        Sheet sheet = workbook.getSheetAt(0);
        List<Map<String,String>> list=new ArrayList<>();
        //创建一个int值，记录当前应该取list的第几个
        int scroll=0;
        // 数据处理
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                //存储结果的map
                if(row.getCell(0)!=null){
                    Map<String,String> resultMap=new HashMap<>();
                    //资源不为null，下一个资源为null，这里面是一个整体
                    resultMap.put("资源名称",row.getCell(0).toString());
                    resultMap.put("会话类型",row.getCell(1).toString());
                    resultMap.put("用户名",row.getCell(2).toString());
                    resultMap.put("账号",row.getCell(3).toString());
                    resultMap.put("服务IP:端口",row.getCell(4).toString());
                    list.add(resultMap);
                    scroll++;
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
                            Map map = list.get(scroll - 1);
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
        FileOutputStream outputStream = new FileOutputStream(new File("E:\\Processed_Data1.xlsx"));
        writeExcelFile(outputStream, list);
        System.out.println("数据处理并写入新的 excel 文件完成！");
    }

    private static void writeExcelFile( FileOutputStream outputStream, List<Map<String, String>> dataList) throws IOException, IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");
        int rowIndex = 1;
        boolean flag=true;
        for (Map<String, String> data : dataList) {
            if(flag){
                //写titile
                int cellIndex0 = 0;
                Row row0 = sheet.createRow(0);
                for (String key : data.keySet()) {
                    if(!key.equals("need")){
                        Cell cell = row0.createCell(cellIndex0++);
                        cell.setCellValue(key);
                    }
                }
                flag=false;
            }


            if(data.containsKey("need")){
                Row row = sheet.createRow(rowIndex++);
                int cellIndex = 0;
                for (String key : data.keySet()) {
                    if(!key.equals("need")){
                        Cell cell = row.createCell(cellIndex++);
                        cell.setCellValue(data.get(key));
                    }
                }
            }
        }
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();

    }
}