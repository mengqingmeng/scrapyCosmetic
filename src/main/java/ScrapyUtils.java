import com.alibaba.fastjson.JSONArray;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * MQM
 * 2018/8/28 17:49
 */
@Slf4j
public class ScrapyUtils {
    private static final String commonUrl = "https://www.cosmetic-info.jp/prod";

    private static final String baseUrl = "https://www.cosmetic-info.jp/prod/search.php";

    private static final String listUrl = "https://www.cosmetic-info.jp/prod/result.php";

    private static File outFile = new File("cosmetic.xlsx");

    private static Workbook wb;
    private static Sheet sheet;
    private static int total = 0;
    private static FileOutputStream stream;

    public static void scrapy() throws IOException {
        if(!outFile.exists()){
            outFile.createNewFile();
        }

        XSSFWorkbook wb =new XSSFWorkbook ();
        sheet = wb.createSheet("第一页");
        Row firstRow = sheet.createRow(0);
        Cell productNameCell = firstRow.createCell(0, CellType.STRING);
        productNameCell.setCellValue("产品名");
        Cell companyCell = firstRow.createCell(1, CellType.STRING);
        companyCell.setCellValue("公司名");
        Cell dateCell = firstRow.createCell(2, CellType.STRING);
        dateCell.setCellValue("日期");
        Cell clazzCeel = firstRow.createCell(3, CellType.STRING);
        clazzCeel.setCellValue("分类");
        Cell cfCeel = firstRow.createCell(4, CellType.STRING);
        cfCeel.setCellValue("成分");
        total ++;
        Document doc = Jsoup.connect(baseUrl).timeout(50000).get();
        Element td = doc.select("td.checkboxArea").first();

        Elements labels = td.getElementsByAttribute("for");
        for(Element label : labels){
            if (label.childNodes().size()>0){
                Node node = label.childNode(0);
                String attr = label.attr("for");
                if(attr.contains("itemCategory")){
                    attr = attr.replace("itemCategory","");
                    int categoryNumber = Integer.parseInt(attr);
                    String nodeStr = node.toString();
                    System.out.println("开始爬："+nodeStr+"类别");
                    scrapySelect(nodeStr,categoryNumber);
                    FileOutputStream stream= FileUtils.openOutputStream(outFile);

                    wb.write(stream);
                    stream.close();
                    System.out.println("爬取："+nodeStr+"类别-完成");
                }
            }
        }

        System.out.println("爬取完成,共："+total--);
    }

    public static void scrapySelect(String select,int categoryNumber){

        Map<String, String> map = new HashMap<String, String>();
        map.put("itemCategoryRowData[]", String.valueOf(categoryNumber));
        map.put("-f", "saler");
        map.put("-d", "a");
        map.put("-p", "all");

        Document doc = null;
        try {
            doc = Jsoup.connect(listUrl).data(map).timeout(50000).post();
        } catch (IOException e) {
            log.error("请求超时："+e.getMessage());
        }
        if(doc == null)
            return;
        Elements elements = doc.getElementsByTag("tbody");
        Element tbody = elements.first();
        Elements allTr = tbody.getElementsByTag("tr");
        for (Element tr:allTr) {
            String productName = "";
            Elements tds = tr.getElementsByTag("td");


            Element productEOuter = tds.get(1);
            Elements productE = productEOuter.select("a.detailicon");

            String detailHref = "";
            if(productE.size() > 0 ){
                Element nameE = productE.get(0);
                if (nameE.hasAttr("href")){
                    detailHref = nameE.attr("href");
                }
                productName = nameE.childNode(0).toString();
//                System.out.println(productName);
            }

//            System.out.println("tds size:"+tds.size());
            String company = "";
            if(tds.size()>2){
                Element companyE = tds.get(2);
                if(companyE.childNodes().size()>0){
                    company = companyE.childNode(0).toString();
                }
            }

            String date = "";
            if(tds.size()>3){
                Element dateE = tds.get(3);
                if (dateE.childNodes().size() > 0) {
                    date = dateE.childNode(0).toString();
                }
            }

            String clazz = "";
            if(tds.size()>4){
                clazz = tds.get(4).childNode(0).toString();
            }


            Row row = sheet.createRow(total);
            Cell productNameCell = row.createCell(0, CellType.STRING);
            productNameCell.setCellValue(productName);
            Cell companyCell = row.createCell(1, CellType.STRING);
            companyCell.setCellValue(company);
            Cell dateCell = row.createCell(2, CellType.STRING);
            dateCell.setCellValue(date);
            Cell clazzCeel = row.createCell(3, CellType.STRING);
            clazzCeel.setCellValue(clazz);

            if(detailHref != null && detailHref.length() > 0){
                String detail = getDetail(detailHref);
                if(detail!=null){
                    Cell detailCeel = row.createCell(4, CellType.STRING);
                    detailCeel.setCellValue(detail);
                }
            }
            total++;

//            System.out.println("productName:"+productName+";company:"+company+";date:"+date+";clazz:"+clazz);
        }
    }

    public static String getDetail(String detailHref){
        String fullUrl = commonUrl+"/"+detailHref;
        try {
            Document doc = Jsoup.connect(fullUrl).timeout(50000).get();
            Elements ols = doc.getElementsByTag("ol");
            if(ols.size() == 1){
                Element ol = ols.get(0);
                Elements as = ol.getElementsByTag("a");
                JSONArray jsonArray = new JSONArray();
                for (Element a :as){
                    if(a.childNodes().size()>0){
                        Node aNode = a.childNodes().get(0);
                        jsonArray.add(aNode.toString());
                    }
                }
                return jsonArray.toJSONString();
            }
        } catch (IOException e) {
            log.error("详情页爬取失败："+e.getMessage());
        }
        return null;
    }
}
