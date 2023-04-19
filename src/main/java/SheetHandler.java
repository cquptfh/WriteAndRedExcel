import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

public   class SheetHandler extends DefaultHandler {
    private SharedStringsTable sharedStringsTable;
    private String currentCellValue;
    private boolean isCellString;

    public SheetHandler(SharedStringsTable sharedStringsTable) {
        this.sharedStringsTable = sharedStringsTable;
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        if ("c".equals(qName)) {
            // 读取单元格的类型
            String cellType = attributes.getValue("t");
            // 判断单元格类型是否为字符串
            isCellString = "s".equals(cellType);
        }
        // 清空当前单元格的值
        currentCellValue = "";
    }

    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        // 处理单元格的值
        if ("v".equals(qName)) {
            if (isCellString) {
                int index = Integer.parseInt(currentCellValue);
                currentCellValue = sharedStringsTable.getItemAt(index).getString();
            }
            System.out.println(currentCellValue);
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        // 读取单元格的值
        currentCellValue += new String(ch, start, length);
    }
}
