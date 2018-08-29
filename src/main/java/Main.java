import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;

/**
 * MQM
 * 2018/8/28 17:49
 */
public class Main {
    public static void main(String[] args) {
        try {
            ScrapyUtils.scrapy();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
