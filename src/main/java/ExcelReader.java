import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.net.URL;

public class ExcelReader {

    public static void main(String[] args) throws Exception {
       new ExcelReader(). readChart2();
    }

    public  void readChart2() throws Exception {


       String path= System.getProperty("user.dir")+"\\src\\main\\java\\pufa.xlsx";
       System.out.println(path+"\\src\\main\\java\\pufa.xlsx");
        File excelFile = new File(path);
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(excelFile));
        //获取第一个sheet
        XSSFSheet sheet = workbook.getSheet(workbook.getSheetName(0));
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        for (XSSFChart chart:drawing.getCharts()){
            CTChart ctChart = chart.getCTChart();
            CTPlotArea plotArea = ctChart.getPlotArea();
            // plotArea.getPieChartList().get(0)
            System.out.println("pie" + plotArea.getPieChartList().size());
            System.out.print("line" + plotArea.getLineChartList().size());
            System.out.print("bar" + plotArea.getBarChartList().size());
            System.out.print("scatter" + plotArea.getScatterChartList().size());
        }
       // XSSFChart chart = drawing.getCharts().get(0);

    }

}
