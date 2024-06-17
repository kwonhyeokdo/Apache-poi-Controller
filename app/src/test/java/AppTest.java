import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.jupiter.api.Test;

import excel.ExcelController;
import excel.FileFormat;
import excel.FileObject;
import excel.ImageFormat;
import excel.ImageObject;

class AppTest {
    @Test
    void excelControllerTest() throws IOException {
        new ExcelController()
                .selectWorksheet(0)
                    .setSheetName("Sheet One")
                    .setDefaultColumnWidthInPixels(50)
                    .setDefaultRowHeightInPixels(40)
                        .mergedRegionAndSelectCell(0, 0, 1, 2)
                            .addText("헤더 1")
                            .setBold(true)
                            .setFontColor(100, 0, 100)
                            .setFontPoints((short)14)
                            .setHorizontalAlignment(HorizontalAlignment.CENTER)
                            .setVerticalAlignment(VerticalAlignment.CENTER)
                            .setCellColor(100, 255, 100)
                        .finishWorkcell()
                        .selectCell(1, 1)
                            .setWidthInPixels(200)
                            .setHeightInPixels(200)
                            .addText("꽁꽁 얼어붙은 한강 위로 고양이가 걸어 갑니다.")
                            .addImage(
                                new ImageObject(
                                    ClassLoader.getSystemClassLoader().getResourceAsStream("sample/cat150x100.jpg").readAllBytes(),
                                    ImageFormat.PICTURE_TYPE_JPEG,
                                    "cat.jpg"
                                )
                            )
                            .addText("it's time to go to bed 오죠-사마")
                            .addFile(
                                new FileObject(
                                    ClassLoader.getSystemClassLoader().getResourceAsStream("sample/TestText.txt").readAllBytes(),
                                    FileFormat.ETC,
                                    "TestText.txt"
                                )
                            )
                            .addText("어떻게 지평좌표계로 고정을 하셨죠?")
                        .finishWorkcell()
                        .selectCell(2, 2)
                            .setWidthInPixels(300)
                            .setHeightInPixels(300)
                            .addFile(
                                new FileObject(
                                    ClassLoader.getSystemClassLoader().getResourceAsStream("sample/TestExcel.xlsx").readAllBytes(),
                                    FileFormat.EXCEL,
                                    "TestExcel.xlsx"
                                )
                            )
                            .addText("너 내 도도도도도도")
                            .addImage(
                                new ImageObject(
                                    ClassLoader.getSystemClassLoader().getResourceAsStream("sample/dog200x200.jpg").readAllBytes(),
                                    ImageFormat.PICTURE_TYPE_JPEG,
                                    "dog.jpg"
                                ),
                                10
                            )
                            .setBorderStyle(BorderStyle.THIN)
                            .setBorderColor(255, 0, 0)
                        .finishWorkcell()
                .finishWorksheet()
                .addSheet()
                .selectWorksheet(1)
                    .setSheetName("Sheet Two")
                    .setColumnWidthInPixels(2, 200)
                    .setRowHeightInPixels(2, 200)
                .finishWorksheet()
            .getByteArrayOutputStreamAndClose()
            .writeTo(new FileOutputStream("sample.xlsx"));
        ;
    }
}
