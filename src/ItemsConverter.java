import org.apache.poi.hssf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;
import org.json.XML;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.util.Optional;

public class ItemsConverter {
    static private final int maxGroup = 99;

    public static void main(String[] args) {
        boolean dataPresent;
        HSSFWorkbook workbook;
        try {

            File file = new File("D:\\Michael\\Git\\res_Blind\\Items.xls");
            FileInputStream fIP = new FileInputStream(file);
            workbook = new HSSFWorkbook(fIP);

            if(file.isFile() && file.exists()) {
                System.out.println("Resource item file open successfully.");
                dataPresent = true;
            } else {
                System.out.println("Error to open resource file.");
                dataPresent = false;
            }
            fIP.close();

        } catch (Exception e) {
            System.out.println("Oops!");
            dataPresent = false;
            workbook = new HSSFWorkbook();
        }
        if (dataPresent) {
            short rownum = 1;
            HSSFRow row;
            HSSFSheet sheet = workbook.getSheetAt(0);
            JSONObject itemsJSON = new JSONObject();
            while (!(sheet.getRow(rownum) == null)) {
                row = sheet.getRow(rownum);
                JSONArray itemJSON = new JSONArray();
                itemJSON.put(0, (int) row.getCell(2).getNumericCellValue());
                itemJSON.put(1, (int) row.getCell(3).getNumericCellValue());
                itemJSON.put(2, (int) row.getCell(4).getNumericCellValue());
                itemJSON.put(3, (int) row.getCell(5).getNumericCellValue());
                itemJSON.put(4, row.getCell(6).toString().replace(".0", ""));
                itemJSON.put(5, row.getCell(7).toString().replace(".0", ""));
                itemJSON.put(6, row.getCell(8).toString().replace(".0", ""));
                itemJSON.put(7, row.getCell(9).toString().replace(".0", ""));
                itemJSON.put(8, row.getCell(10).getNumericCellValue());
                itemJSON.put(9, (int) row.getCell(11).getNumericCellValue());
                itemJSON.put(10, row.getCell(12).toString());
                itemJSON.put(11, row.getCell(13).toString().replace(".0", ""));
                itemJSON.put(12, row.getCell(14).getStringCellValue());
                itemJSON.put(13, (int) row.getCell(15).getNumericCellValue());
                itemJSON.put(14, (int) row.getCell(16).getNumericCellValue());
                itemJSON.put(15, (int) row.getCell(17).getNumericCellValue());
                itemJSON.put(16, (int) row.getCell(18).getNumericCellValue());
                itemJSON.put(17, (int) row.getCell(19).getNumericCellValue());
                itemsJSON.put(String.format("%d", (int) row.getCell(0).getNumericCellValue()), itemJSON);
                System.out.println(String.valueOf(rownum));
                rownum++;
            }
            //запись характеристик предметов в файл
            try (FileWriter file = new FileWriter("D:\\Michael\\Git\\res_Blind\\ITEMS.json")) {
                file.write(itemsJSON.toString());
                //file.close();
            } catch (Exception e) {
                System.out.println("Oops!");
            }
            convertItemsStrings(0, sheet);
            convertItemsStrings(1, sheet);
            convertItemsStrings(2, sheet);
        }
    }

    private static void convertItemsStrings (int localID, HSSFSheet sheet) {
        String localStr;
        switch (localID) {
            case 0: localStr = "ru";
                break;
            case 1: localStr = "ua";
                break;
            case 2: localStr = "en";
                break;
            default: System.out.println("File not saved!");
                return;
        }
        JSONArray arrayS;
        JSONArray arrayP;
        JSONArray arrayDesc;
        JSONObject itemsJSON = new JSONObject();
        int rownum;
        HSSFRow row;
        for (int i=0; i<10; i++) {
            arrayS = new JSONArray();
            arrayP = new JSONArray();
            arrayDesc = new JSONArray();
            for (int j=0; j<maxGroup+1; j++) {
                arrayS.put(j, "");
                arrayP.put(j, "");
                arrayDesc.put(j, "");
            }
            rownum = 1;
            while (!(sheet.getRow(rownum) == null)) {
                row = sheet.getRow(rownum);
                int itemID = (int) Math.round(row.getCell(0).getNumericCellValue());
                if ( itemID / 100 == i)  {
                    Optional<HSSFCell> cell = Optional.ofNullable(row.getCell(22 + (localID*3)));
                    arrayS.put(itemID % 100, (cell.isPresent() ? cell.get().getStringCellValue() : ""));
                    cell = Optional.ofNullable(row.getCell(23 + (localID*3)));
                    arrayP.put(itemID % 100, (cell.isPresent() ? cell.get().getStringCellValue() : ""));
                    cell = Optional.ofNullable(row.getCell(24 + (localID*3)));
                    arrayDesc.put(itemID % 100, (cell.isPresent() ? cell.get().getStringCellValue() : ""));
                }
                rownum++;
            }
            //очистка лишних пустых элементов
            int k = maxGroup;
            while (arrayS.getString(k).equals("")) {
                arrayS.remove(k);
                k--;
                if (k < 0) break;
            }
            k = maxGroup;
            while (arrayP.getString(k).equals("")) {
                arrayP.remove(k);
                k--;
                if (k < 0) break;
            }
            k = maxGroup;
            while (arrayDesc.getString(k).equals("")) {
                arrayDesc.remove(k);
                k--;
                if (k < 0) break;
            }
            itemsJSON.put(localStr + "_itemsS" + String.valueOf(i), (new JSONObject()).put("item", arrayS));
            itemsJSON.put(localStr + "_itemsP" + String.valueOf(i), (new JSONObject()).put("item", arrayP));
            itemsJSON.put(localStr + "_itemsDesc" + String.valueOf(i), (new JSONObject()).put("item", arrayDesc));
        }
        String itemsXML = XML.toString(itemsJSON, "resources");
        itemsXML = itemsXML.replaceAll("<(\\w{2})_itemsS(\\d+)>", "<string-array name=\"$1_itemsS$2\">");
        itemsXML = itemsXML.replaceAll("</(\\w{2})_itemsS(\\d+)>", "</string-array>");
        itemsXML = itemsXML.replaceAll("<(\\w{2})_itemsP(\\d+)>", "<string-array name=\"$1_itemsP$2\">");
        itemsXML = itemsXML.replaceAll("</(\\w{2})_itemsP(\\d+)>", "</string-array>");
        itemsXML = itemsXML.replaceAll("<(\\w{2})_itemsDesc(\\d+)>", "<string-array name=\"$1_itemsDesc$2\">");
        itemsXML = itemsXML.replaceAll("</(\\w{2})_itemsDesc(\\d+)>", "</string-array>");

        try (FileWriter file = new FileWriter("D:\\Michael\\Git\\res_Blind\\" + localStr + "_items.xml")) {
            file.write(itemsXML);
            //file.close();
        } catch (Exception e) {
            System.out.println("Oops!");
        }
    }

}
