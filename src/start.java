import com.drew.imaging.ImageMetadataReader;
import com.drew.imaging.ImageProcessingException;
import com.drew.metadata.exif.GpsDirectory;
import com.drew.lang.GeoLocation;
import com.drew.metadata.Metadata;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Collection;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;

public class start {

    public static void main(String[] args) throws IOException, ImageProcessingException {

        //Create Excel objects, sheet and columns
        Workbook wb = new HSSFWorkbook();
        CreationHelper helper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("new sheet");
        Drawing drawing = sheet.createDrawingPatriarch();
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue(helper.createRichTextString("Img"));
        row.createCell(1).setCellValue(helper.createRichTextString("Name"));
        row.createCell(2).setCellValue(helper.createRichTextString("Lat"));
        row.createCell(3).setCellValue(helper.createRichTextString("Lon"));
        sheet.setColumnWidth(0, 12500);
        CellStyle style = wb.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        //Read folder
        File folder = new File("imgs/");
        File[] listOfFiles = folder.listFiles();
        int rowNum = 1;
        for (File file : listOfFiles) {
            if (file.isFile()) {

                // Read image
                FileInputStream stream = new FileInputStream(file);
                byte[] bytes = IOUtils.toByteArray(stream);

                int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
                stream.close();

                // Read the metadata
                double lat = 0;
                double lon = 0;
                Metadata metadata = ImageMetadataReader.readMetadata(file);
                Collection<GpsDirectory> gpsDirectories = metadata.getDirectoriesOfType(GpsDirectory.class);
                for (GpsDirectory gpsDirectory : gpsDirectories) {
                    GeoLocation exifLocation = gpsDirectory.getGeoLocation();
                    lat = exifLocation.getLatitude();
                    lon = exifLocation.getLongitude();
                }

                // Write to Excel sheet
                Row rowVal = sheet.createRow(rowNum);

                // Image
                ClientAnchor anchor = helper.createClientAnchor();
                anchor.setCol1(0);
                anchor.setRow1(rowNum);
                Picture pict = drawing.createPicture(anchor, pictureIdx);
                rowVal.setHeight((short) 4000);
                pict.resize();

                // Image name
                Cell cell = rowVal.createCell(1);
                cell.setCellValue(helper.createRichTextString(file.getName()));
                cell.setCellStyle(style);

                // Lat
                cell = rowVal.createCell(2);
                cell.setCellValue(helper.createRichTextString(Double.toString(lat)));
                cell.setCellStyle(style);

                //Lon
                cell = rowVal.createCell(3);
                cell.setCellValue(helper.createRichTextString(Double.toString(lon)));
                cell.setCellStyle(style);

                rowNum++;
            }
        }

        // Adjust column sizes
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);

        // Write the Excel file to disk
        try (OutputStream fileOut = new FileOutputStream("table.xls")) {
            wb.write(fileOut);
        }

    }

    private static String getFileExtension(File file) {
        String extension = "";

        try {
            if (file != null && file.exists()) {
                String name = file.getName();
                extension = name.substring(name.lastIndexOf("."));
            }
        } catch (Exception e) {
            extension = "";
        }

        return extension;

    }
}
