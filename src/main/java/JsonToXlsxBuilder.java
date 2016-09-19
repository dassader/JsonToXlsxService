import org.apache.poi.ss.usermodel.Workbook;

import java.io.InputStream;

/**
 * Created by AndriiK on 9/19/2016.
 */
public interface JsonToXlsxBuilder {
    InputStream build(String jsonData);
}
