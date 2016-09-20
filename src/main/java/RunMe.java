import org.apache.commons.io.IOUtils;

import java.io.*;


/**
 * Created by AndriiK on 9/19/2016.
 */
public class RunMe {
    public static void main(String[] args) throws IOException {
        String testJsonData = readFile("input.json");

        JsonToXlsxBuilder jsonToXlsxBuilder = new PoiJsonToXlsxBuilder();
        InputStream data = jsonToXlsxBuilder.build(testJsonData);

        IOUtils.copy(data, new FileOutputStream("output.xlsx"));

        String program = "C:\\Program Files (x86)\\LibreOffice 5\\program\\scalc.exe";

        File file = new File("output.xlsx");

        Runtime.getRuntime().exec(program+" "+file.getAbsolutePath());
    }

    public static String readFile(String name) throws IOException {
        InputStream in = new FileInputStream(name);

        try {
            return IOUtils.toString(in);
        } finally {
            IOUtils.closeQuietly(in);
        }
    }
}
