import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.commons.io.IOUtils;

import java.io.*;
import java.util.List;


public class RunMe {
    public static void main(String[] args) throws IOException {
        String testJsonData = readFile("input3.json");

        ObjectMapper objectMapper = new ObjectMapper();

        JsonToXlsxBuilder jsonToXlsxBuilder = new PoiJsonToExcelBuilder();
        List<JsonSheet> jsonSheets = objectMapper.readValue(testJsonData, new TypeReference<List<JsonSheet>>(){});
        byte[] butes = jsonToXlsxBuilder.build(jsonSheets);

        IOUtils.copy(new ByteArrayInputStream(butes), new FileOutputStream("output.xlsx"));

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
