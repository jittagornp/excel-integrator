/*
 *  Copy right 2016 pamarin.com
 */
import java.io.File;
import java.io.IOException;
import java.net.URL;
import org.junit.Test;
import com.pamarin.util.excel.integrator.ExcelFile;
import com.pamarin.util.excel.integrator.ExcelSheetIntegrator;

/**
 * @author jittagornp
 */
public class ExcelSheetIntegratorTest {

    @Test
    public void test() throws IOException {

        URL url = getClass().getResource("/excel");
        String path = url.getPath();

        System.out.println(path);

        File input1 = new File(path, "file1.xlsx");
        File input2 = new File(path, "file2.xlsx");
        File input3 = new File(path, "file3.xlsx");

        ExcelFile exFile1 = ExcelFile.from(input1).andWithSheetName("ชื่อ sheet 1").andWithSheetName("ชื่อ sheet 2");
        ExcelFile exFile2 = ExcelFile.from(input2).andWithSheetName("ชื่อ sheet 3");
        ExcelFile exFile3 = ExcelFile.from(input3).andWithSheetName("ชื่อ sheet 4");

        File output1 = new File(path, "output.xlsx");

        File integratedFile = ExcelSheetIntegrator.newInstance()
                .addExcelFile(exFile1)
                .addExcelFile(exFile2)
                .addExcelFile(exFile3)
                .toTargetFile(output1)
                .integrate();
    }

}
