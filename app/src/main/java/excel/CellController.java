package excel;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.ObjectData;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDvAspect;

public class CellController {
    private final int rowIndex;
    private final int colIndex;
    private final ExcelController excelController;
    private final SheetController sheetController;
    private final Workbook workbook;
    private final Sheet worksheet;
    private final Cell workcell;
    private final Row workrow;
    private CellStyle workcellStyle;
    private Font workfont;
    private final Map<Integer, Picture> pictureList = new HashMap<>();

    public CellController(final SheetController sheetController, final int rowIndex, final int colIndex) {
        this.rowIndex = rowIndex;
        this.colIndex = colIndex;
        this.sheetController = sheetController;
        this.excelController = sheetController.getExcelController();
        this.workbook = excelController.getWorkbook();
        this.worksheet = sheetController.getWorksheet();

        this.workcellStyle = workbook.createCellStyle();
        workcellStyle.setVerticalAlignment(VerticalAlignment.TOP); // 글자 위쪽 맞춤
        workcellStyle.setWrapText(true); // 텍스트 줄 바꿈

        workrow = sheetController.getRow(rowIndex);
        workcell = workrow.createCell(colIndex);
        workcell.setCellStyle(workcellStyle);
    }

    /**
     * 내용의 수직 정렬을 설정한다.
     * @param verticalAlignment
     * @return this
     */
    public CellController setVerticalAlignment(final VerticalAlignment verticalAlignment){
        workcellStyle.setVerticalAlignment(verticalAlignment);
        return this;
    }

    /**
     * 내용의 수평 정렬을 설정한다.
     * @param horizontalAlignment
     * @return this
     */
    public CellController setHorizontalAlignment(final HorizontalAlignment horizontalAlignment){
        workcellStyle.setAlignment(horizontalAlignment);
        return this;
    }

    /**
     * Cell(Column)의 Width를 width으로 변경한다.
     * @param pixels 픽셀
     * @return this
     */
    public CellController setWidth(final int width){
        sheetController.setColumnWidth(colIndex, width);
        return this;
    }

    /**
     * Cell(Column)의 Width를 pixel으로 변경한다.
     * @param pixels 픽셀
     * @return this
     */
    public CellController setWidthInPixels(final int pixels){
        sheetController.setColumnWidthInPixels(colIndex, pixels);
        return this;
    }

    /**
     * Cell(Row)의 Height를 points으로 변경한다.
     * @param pixels
     * @return this
     */
    public CellController setHeightInPoints(final int points){
        sheetController.setRowHeightInPoints(rowIndex, points);
        return this;
    }

    /**
     * Cell(Row)의 Height를 pixels으로 변경한다.
     * @param pixels
     * @return this
     */
    public CellController setHeightInPixels(final int pixels){
        sheetController.setRowHeightInPixels(rowIndex, pixels);
        return this;
    }

    /**
     * Cell의 Text를 변경한다.
     * @param text 입력할 Text
     * @return this
     */
    public CellController setText(final String text){
        workcell.setCellValue(text);
        return this;
    }

    /**
     * Cell의 Number를 변경한다.
     * @param value 입력할 value
     * @return this
     */
    public CellController setNumber(final int value){
        workcell.setCellValue(value);
        return this;
    }

    /**
     * Cell의 Number를 변경한다.
     * @param value 입력할 value
     * @return this
     */
    public CellController setNumber(final float value){
        workcell.setCellValue(value);
        return this;
    }

    /**
     * Cell의 Number를 변경한다.
     * @param value 입력할 value
     * @return this
     */
    public CellController setNumber(final double value){
        workcell.setCellValue(value);
        return this;
    }

    /**
     * Cell에 설정된 workcellStyle 인스턴스를 반환한다.
     * @return workcellStyle
     */
    public CellStyle getWorkcellStyle(){
        return workcellStyle;
    }

    /**
     * workcell을 cellStyle로 교체한다.
     * @param cellStyle
     * @return this
     */
    public CellController setCellStyle(final CellStyle cellStyle){
        workcellStyle = cellStyle;
        workcell.setCellStyle(workcellStyle);
        return this;
    }


    /**
     * R, G, B로 XSSFColor를 생성후 반환한다.
     * workbook 구현체가 XSSFWorkbook이 아닐경우 null을 반환한다.
     * @param R
     * @param G
     * @param B
     * @return R, G, B로 XSSFColor를 생성 후 Color를 반환한다.
     */
    private Color getColor(final int R, final int G, final int B){
        if(workbook instanceof XSSFWorkbook){
            XSSFWorkbook xssfWorkbook = (XSSFWorkbook)workbook;
            IndexedColorMap indexedColors = xssfWorkbook.getStylesSource().getIndexedColors();
            return new XSSFColor(new java.awt.Color(R, G, B), indexedColors);
        }else{
            return null;
        }
    }

    /**
     * Cell의 색상을 변경한다.
     * @param R
     * @param G
     * @param B
     * @return this
     */
    public CellController setCellColor(final int R, final int G, final int B){
        Color color = getColor(R, G, B);

        workcellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        workcellStyle.setFillForegroundColor(color);

        return this;
    }

    /**
     * Cell의 Top-Border Style를 설정한다.
     * @param borderStyle Border의 스타일
     * @return this
     */
    public CellController setTopBorderStyle(final BorderStyle borderStyle){
        workcellStyle.setBorderTop(borderStyle);
        return this;
    }

    /**
     * Cell의 Bottom-Border Style를 설정한다.
     * @param borderStyle Border의 스타일
     * @return this
     */
    public CellController setBottomBorderStyle(final BorderStyle borderStyle){
        workcellStyle.setBorderBottom(borderStyle);
        return this;
    }

    /**
     * Cell의 Left-Border Style를 설정한다.
     * @param borderStyle Border의 스타일
     * @return this
     */
    public CellController setLeftBorderStyle(final BorderStyle borderStyle){
        workcellStyle.setBorderLeft(borderStyle);
        return this;
    }

    /**
     * Cell의 Right-Border를 설정한다.
     * @param borderStyle Border의 스타일
     * @return this
     */
    public CellController setRightBorderStyle(final BorderStyle borderStyle){
        workcellStyle.setBorderRight(borderStyle);
        return this;
    }

    /**
     * Cell의 Border Style를 설정한다.
     * @param borderStyle Border의 스타일
     * @return this
     */
    public CellController setBorderStyle(final BorderStyle borderStyle){
        return this
              .setTopBorderStyle(borderStyle)
              .setBottomBorderStyle(borderStyle)
              .setLeftBorderStyle(borderStyle)
              .setRightBorderStyle(borderStyle)
        ;
    }

    /**
     * Cell의 Top-Border Color를 설정한다.
     * workbook의 구현체가 XSSFWorkbook이어야 적용된다.
     * @param R
     * @param G
     * @param B
     * @return this
     */
    public CellController setTopBorderColor(final int R, final int G, final int B){
        if(workbook instanceof XSSFWorkbook){
            XSSFColor color = (XSSFColor)getColor(R, G, B);
            XSSFCellStyle xssfCellStyle = (XSSFCellStyle)workcellStyle;
            xssfCellStyle.setTopBorderColor(color);
        }
        return this;
    }

    /**
     * Cell의 Bottom-Border Color를 설정한다.
     * workbook의 구현체가 XSSFWorkbook이어야 적용된다.
     * @param R
     * @param G
     * @param B
     * @return this
     */
    public CellController setBottomBorderColor(final int R, final int G, final int B){
        if(workbook instanceof XSSFWorkbook){
            XSSFColor color = (XSSFColor)getColor(R, G, B);
            XSSFCellStyle xssfCellStyle = (XSSFCellStyle)workcellStyle;
            xssfCellStyle.setBottomBorderColor(color);
        }
        return this;
    }

    /**
     * Cell의 LEFT-Border Color를 설정한다.
     * workbook의 구현체가 XSSFWorkbook이어야 적용된다.
     * @param R
     * @param G
     * @param B
     * @return this
     */
    public CellController setLeftBorderColor(final int R, final int G, final int B){
        if(workbook instanceof XSSFWorkbook){
            XSSFColor color = (XSSFColor)getColor(R, G, B);
            XSSFCellStyle xssfCellStyle = (XSSFCellStyle)workcellStyle;
            xssfCellStyle.setLeftBorderColor(color);
        }
        return this;
    }

    /**
     * Cell의 Right-Border Color를 설정한다.
     * workbook의 구현체가 XSSFWorkbook이어야 적용된다.
     * @param R
     * @param G
     * @param B
     * @return this
     */
    public CellController setRightBorderColor(final int R, final int G, final int B){
        if(workbook instanceof XSSFWorkbook){
            XSSFColor color = (XSSFColor)getColor(R, G, B);
            XSSFCellStyle xssfCellStyle = (XSSFCellStyle)workcellStyle;
            xssfCellStyle.setRightBorderColor(color);
        }
        return this;
    }

    /**
     * Cell의 Border Color를 설정한다.
     * workbook의 구현체가 XSSFWorkbook이어야 적용된다.
     * @param R
     * @param G
     * @param B
     * @return this
     */
    public CellController setBorderColor(final int R, final int G, final int B){
        return this
              .setTopBorderColor(R, G, B)
              .setBottomBorderColor(R, G, B)
              .setLeftBorderColor(R, G, B)
              .setRightBorderColor(R, G, B)
        ;
    }

    /**
     * Cell에 삽입한 Image의 테두리 색을 설정한다.
     * @param String getImageKeySet()으로 확인할 수 있다.
     * @param R
     * @param G
     * @param B
     * @return this
     */
    public CellController setImageLineColor(
        final String imageKey,
        final int R, final int G, final int B
    ){
        if(getImageKeySet().contains(imageKey)){
            int imageIndex = excelController.getImageIndexMap().get(imageKey);
            pictureList.get(imageIndex).setLineStyleColor(R, G, B);
        }
        return this;
    }

    /**
     * workFont를 반환한다.
     * workFont가 없을 경우 새로 생성 후 반환한다.
     * @return workFont를 반환한다.
     */
    private Font getWorkFont(){
        if(workfont == null){
            workfont = workbook.createFont();
            workfont.setFontName(Base.BASE_FONT_NAME);
            workfont.setFontHeightInPoints(Base.BASE_FONT_HEIGHT_POINTS);
            workcellStyle.setFont(workfont);
        }
        return workfont;
    }

    /**
     * Cell의 Font Points 설정한다.
     * @param points 폰트 크기
     * @return this
     */
    public CellController setFontPoints(final short points){
        getWorkFont().setFontHeightInPoints(points);
        return this;
    }

    /**
     * Cell의 font color를 설정한다.
     * workbook의 구현체가 XSSFWorkbook이어야 적용된다.
     * @param R
     * @param G
     * @param B
     * @return this
     */
    public CellController setFontColor(int R, int G, int B){
        if(workbook instanceof XSSFWorkbook){
            XSSFColor color = (XSSFColor)getColor(R, G, B);
            XSSFFont xssfFont = (XSSFFont)getWorkFont();
            xssfFont.setColor(color);
        }
        return this;
    }

    /**
     * Font의 Bold 여부
     * @param bold
     * @return this
     */
    public CellController setBold(boolean bold){
        getWorkFont().setBold(bold);
        return this;
    }

    /**
     * Cell의 DataFormat을 설정한다.
     * 참고 [표현형식 Index] - https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/BuiltinFormats.html
     * @param dataformatIndex 표현형식 Index, 예시) 0x31 : text
     * @return this
     */
    public CellController setDataFormat(int dataformatIndex){
        if(workbook instanceof XSSFWorkbook){
            XSSFCellStyle xssfCellStyle = (XSSFCellStyle)workcellStyle;
            xssfCellStyle.setDataFormat(dataformatIndex);
        }else{
            workcellStyle.setDataFormat((short)dataformatIndex);
        }
        return this;
    }

    /**
     * Cell의 DataFormat을 설정한다.
     * workbook의 구현체가 XSSFWorkbook이어야 적용된다.
     * 참고 [표현형식 Index] - https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/BuiltinFormats.html
     * @param dataformat 표현형식 Index, 예시) "#,##0"
     * @return 현재 인스턴스(CellController)
     */
    public CellController setDataFormat(String dataformat){
        if(workbook instanceof XSSFWorkbook){
            XSSFCellStyle xssfCellStyle = (XSSFCellStyle)workcellStyle;
            xssfCellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(dataformat));
        }
        return this;
    }

    /**
     * Image 처리를 위해 사용되는 등록된 ImageKey들을 Set형태로 리턴한다.
     * @return Image 처리를 위해 사용되는 등록된 ImageKey들을 Set형태로 리턴한다.
     */
    public Set<String> getImageKeySet(){
        return excelController.getImageKeySet();
    }

    /**
     **<pre>
     **1. Cell에 Image를 넣는다.
     **2. imageObject의 imageKey가 기존에 사용/등록 되었다면, 기존 Image byte[]를 사용한다. 따라서 기존에 등록된 imageKey를 넣으면 기존 이미지를 사용할 수 있고, 새로운 Image를 사용하기 위해서는 imageKey 값이 중복되지 않게 설정해야 한다.
     **3. 기존의 등록된 ImageKey는 getImageKeySet()로 확인한다.
     **4. positionObject의 dx, dy의 기준은 px이다.
     **5. Cell의 Width와 Height를 초과하도록 px가 설정되어도 Image의 실제 크기는 Width와 Height 보다 클 수 없다.
     **6. 참고 - https://stackoverflow.com/questions/47503477/apache-poi-write-image-and-text-excel
     * </pre>
     * @param imageObject
     * @param positionObject
     * @return this
     */
    public CellController setImage(final ImageObject imageObject, final Position positionObject){
        final String imageKey = imageObject.getImageKey();
        final byte[] imageByteArray = imageObject.getImageByteArray();
        final ImageFormat imageFormat = imageObject.getImageFormat();
        final int dx1 = positionObject.getDx1();
        final int dy1 = positionObject.getDy1();
        final int dx2 = positionObject.getDx2();
        final int dy2 = positionObject.getDy2();

        int imageIndex = -99999;
        final Map<String, Integer> imageIndexMap = excelController.getImageIndexMap();
        if(imageIndexMap.containsKey(imageKey)){
            imageIndex = imageIndexMap.get(imageKey);
        }else{
            imageIndex = workbook.addPicture(imageByteArray, imageFormat.getValue());
            imageIndexMap.put(imageKey, imageIndex);
        }

        XSSFClientAnchor anchor = new XSSFClientAnchor();
        anchor.setRow1(rowIndex);
        anchor.setRow2(rowIndex);
        anchor.setCol1(colIndex);
        anchor.setCol2(colIndex);
        anchor.setDx1(Units.EMU_PER_PIXEL * dx1);
        anchor.setDy1(Units.EMU_PER_PIXEL * dy1);
        anchor.setDx2(Units.EMU_PER_PIXEL * dx2);
        anchor.setDy2(Units.EMU_PER_PIXEL * dy2);
        anchor.setAnchorType(AnchorType.MOVE_DONT_RESIZE);

        Picture picture = sheetController.getWorkdrawing().createPicture(anchor, imageIndex);
        pictureList.put(imageIndex, picture);

        return this;
    }

    /**
     **<pre>
     **1. Cell에 Embedded File를 넣는다.
     **2. file fileName 기존에 사용/등록 되었다면, 기존 fileByteArray를 사용한다. 따라서 기존에 등록/사용된 fileName를 넣으면 기존 file을 사용할 수 있고, 새로운 file를 사용하기 위해서는 fileName 값이 중복되지 않게 설정해야 한다.
     **3. 기존의 등록된 fileName getfileNameSet()로 확인한다.
     **4. positionObject의 dx, dy의 기준은 px이다.
     **5. Cell의 Width와 Height를 초과하도록 px가 설정되어도 파일 아이콘 Image의 실제 크기는 Width와 Height 보다 클 수 없다.
     * </pre>
     * @param file
     * @param position
     * @return this
     * @throws IOException
     */
    public CellController setFile(final FileObject fileObject, final Position position) throws IOException{
        final String fileName = fileObject.getFileName();
        final byte[] fileByteArray = fileObject.getFileByteArray();
        final FileFormat fileFormat = fileObject.getFileFormat();
        final int dx1 = position.getDx1();
        final int dy1 = position.getDy1();
        final int dx2 = position.getDx2();
        final int dy2 = position.getDy2();

        int fileIndex = -99999;
        final Map<String, Integer> imageIndexMap = excelController.getImageIndexMap();
        final Map<String, Integer> fileIndexMap = excelController.getFileIndexMap();
        if(fileIndexMap.containsKey(fileName)){
            fileIndex = fileIndexMap.get(fileName);
        }else{
            fileIndex = workbook.addOlePackage(fileByteArray, fileName, fileName, fileName);
            fileIndexMap.put(fileName, fileIndex);
        }
        int imageIndex = imageIndexMap.get(fileFormat.getIconName());

        XSSFClientAnchor anchor = new XSSFClientAnchor();
        anchor.setRow1(rowIndex);
        anchor.setRow2(rowIndex);
        anchor.setCol1(colIndex);
        anchor.setCol2(colIndex);
        anchor.setDx1(Units.EMU_PER_PIXEL * dx1);
        anchor.setDy1(Units.EMU_PER_PIXEL * dy1);
        anchor.setDx2(Units.EMU_PER_PIXEL * dx2);
        anchor.setDy2(Units.EMU_PER_PIXEL * dy2);
        
        ObjectData objectData = sheetController.getWorkdrawing().createObjectData(anchor, fileIndex, imageIndex);
        if(objectData instanceof XSSFObjectData){
            XSSFObjectData xSSFObjectData = (XSSFObjectData)objectData;
            xSSFObjectData.getOleObject().setDvAspect(STDvAspect.DVASPECT_ICON); // 파일 이미지를 더블클릭 했을 때, 엑셀 기능에 의해 썸네일 형식으로 전환되는 것을 방지.
        }

        return this;
    }


    /**
     * 한 Line에 몇 글자가 들어갈 수 있는지 판단하여 text의 총 Line 수를 구한다.
     * @param text
     * @return 한 Line에 몇 글자가 들어갈 수 있는지 판단하여 text의 총 Line 수를 구한다.
     */
    private int getLineCountFromText(CharSequence text, double maxCharacterCountInWidth){
        int newLineCnt = 0;
        double textCnt = 0d;

        for(int i = 0; i < text.length(); i++){
            char c = text.charAt(i);

            if(c == '\n' || c == '\r'){
                if(textCnt > 0d){
                    double lineCnt = textCnt / maxCharacterCountInWidth;
                    newLineCnt += (int)(Math.ceil(lineCnt));
                    textCnt = 0d;
                }else{
                    newLineCnt++;
                }
                continue;
            }else if(c == '"' || c == '\'' || c == '.' || c == ','){
                // continue;
            }else if(c == 'l' || c == 'i' || c == 'j'){
                textCnt += 0.25;
            }else if(c == '(' || c == ')' || c == '{' || c == '}' || c == '[' || c == ']' || c == '!' || c == 'f' || c == 't' || c == 'I'){
                textCnt += 0.3333;
            }
            else if(c == ' ' || c == '-' || c == '_' || c == '*' ||  Character.isDigit(c) || (c >= 'a' && c <= 'z')){
                textCnt += 0.5;
            }else if(c >= 'A' && c <= 'Z'){
                textCnt += 0.8;
            }else{
                textCnt += 1d;
            }

            if(i == text.length() - 1){
                double lineCnt = textCnt / maxCharacterCountInWidth;
                newLineCnt += (int)(Math.ceil(lineCnt));
            }
        }

        return newLineCnt;
    }

    /**
     * 글자의 높이를 Pixels로 구한다.
     * Font Points를 Pixels로 변환한 후 + a값을 더한다.
     * a 값이 무엇인지 정확하지가 않다. 맑은고딕 10pt 기준으로 4이다.
     * @return 글자의 높이를 Pixel로 구한다.
     */
    private int getFontHeightPixels(){
        final int fontPoints = getWorkFont().getFontHeightInPoints();
        return Base.getCharacterHeightPixels(fontPoints);
    }

    /**
     * text가 한줄 또는 여러줄 일 경우 높이가 몇 Pixels인지 구한다.
     * @param text
     * @return text가 한줄 또는 여러줄 일 경우 높이가 몇 Pixels인지 구한다.
     */
    private int getTextHeightPixels(final String text){
        final int fontPoints = getWorkFont().getFontHeightInPoints();
        final int cellWidth = worksheet.getColumnWidth(colIndex);
        final int cellWidthPixels = UnitConverter.columnWidthToPixels(cellWidth);
        final int fontPixels = UnitConverter.pointsToPixels(fontPoints);
        final int fontHeightPixels = getFontHeightPixels();
        final double maxCharacterCountInWidth = cellWidthPixels / fontPixels;
        
        final int lineCnt = getLineCountFromText(text, maxCharacterCountInWidth);
        final int textHeightPixels = lineCnt * fontHeightPixels;

        return textHeightPixels;
    }

    /**
     * heightPixels이 몇 Line인지 구한다.
     * @param heightPixels
     * @return heightPixel이 몇 Line인지 구한다.
     */
    private int getLineCountFromHeightPixels(final int heightPixels){
        final int fontHeightPixel = getFontHeightPixels();
        return (int)(Math.ceil((double)heightPixels / fontHeightPixel));
    }

    /**
     * 텍스트를 이어서 추가한다.
     * @param text
     * @return this
     */
    public CellController addText(final String text){
        if(text != null && text.length() > 0){
            int maxHeightPixels = UnitConverter.pointsToPixels(workrow.getHeightInPoints());
        
            String fullText = workcell.getStringCellValue() + text;
            
            int textHeightPixels = getTextHeightPixels(fullText);
            
            if(maxHeightPixels < textHeightPixels){
                maxHeightPixels = textHeightPixels;
            }

            setText(fullText);
            setHeightInPixels(maxHeightPixels);
        }

        return this;
    }

    /**
     * 이미지를 추가한다.
     * 이미지 크기는 셀의 넓이를 넘어서지 못한다.
     * 이미지 크기는 원본 이미지의 크기를 넘어서지 못한다.
     * 이미지의 가로 세로 크기 중 더 큰 것을 기준으로 셀의 넓이에 따라 사이즈가 조정된다.
     * @param imageObject
     * @return this
     */
    public CellController addImage(final ImageObject imageObject){
        return addImage(imageObject, 0);
    }

    /**
     * 이미지를 추가한다.
     * 이미지 크기는 셀의 넓이를 넘어서지 못한다.
     * 이미지 크기는 원본 이미지의 크기를 넘어서지 못한다.
     * 이미지의 가로 세로 크기 중 더 큰 것을 기준으로 셀의 넓이에 따라 사이즈가 조정된다.
     * @param imageObject
     * @param padding
     * @return this
     */
    public CellController addImage(final ImageObject imageObject, final int padding){
        if(imageObject != null){
            final int cellWidth = worksheet.getColumnWidth(colIndex);
            final int cellWidthPixels = UnitConverter.columnWidthToPixels(cellWidth);
            final int maxHeightPixels = UnitConverter.pointsToPixels(workrow.getHeightInPoints());
            final int sourceImageWidth = imageObject.getWidth();
            final int sourceImageHeight = imageObject.getHeight();
            
            double scale = 1d;
            int imageWidthPixels = 0;
            int imageHeightPixels = 0;

            if(sourceImageWidth < cellWidthPixels && sourceImageHeight < cellWidthPixels){
                imageWidthPixels = sourceImageWidth;
                imageHeightPixels = sourceImageHeight;
            }else{
                if(sourceImageWidth > sourceImageHeight){
                    scale = (double)cellWidthPixels / sourceImageWidth;
                }else{
                    scale = (double)cellWidthPixels / sourceImageHeight;
                }
                imageWidthPixels = (int)(sourceImageWidth * scale);
                imageHeightPixels = (int)(sourceImageHeight * scale);
            }

            int imageLineCount = getLineCountFromHeightPixels(imageHeightPixels);

            String text = workcell.getStringCellValue();
            int fromTextHeightPixels = getTextHeightPixels(text);

            boolean isStart = text.length() == 0;
            for(int i = (isStart ? 1 : 0); i < imageLineCount + 1; i++){ // 엑셀에서 기본으로 빈 텍스트는 한 줄임.
                text += "\n";
            }

            int toTextHeightPixels = getTextHeightPixels(text);

            if(maxHeightPixels < toTextHeightPixels){
                setHeightInPixels(toTextHeightPixels);
            }

            setImage(imageObject, new Position(0 + padding, fromTextHeightPixels + padding, imageWidthPixels - padding, fromTextHeightPixels + imageHeightPixels - padding));

            setText(text);
        }

        return this;
    }

    /**
     * 파일을 추가한다.
     * 아이콘의 크기는 30 x 30 Pixel이다.
     * @param file
     * @return this
     * @throws IOException
     */
    public CellController addFile(final FileObject file) throws IOException{
        return addFile(file, 0);
    }

    /**
     * 파일을 추가한다.
     * 아이콘의 크기는 30 x 30 Pixel이다.
     * @param file
     * @return 현재 인스턴스(CellController)
     * @throws IOException
     */
    public CellController addFile(final FileObject file, final int padding) throws IOException{
        if(file != null){
            final int maxHeightPixel = UnitConverter.pointsToPixels(workrow.getHeightInPoints());
            final int size = 30;

            int imageLineCount = getLineCountFromHeightPixels(size);

            String text = workcell.getStringCellValue();
            boolean isStart =
                (text.length() == 0) ||
                (
                    text.length() >= 2 &&
                    (
                        (text.charAt(text.length() - 1) == '\n' && text.charAt(text.length() - 2) == '\n') ||
                        (text.charAt(text.length() - 1) == '\r' && text.charAt(text.length() - 2) == '\r')
                    )
                )
            ;

            int fromTextHeightPixel = getTextHeightPixels(text);

            for(int i = (isStart ? 1 : 0); i < imageLineCount + 1; i++){ // 엑셀에서 기본으로 빈 텍스트는 한 줄임.
                text += "\n";
            }

            int toTextHeightPixel = getTextHeightPixels(text);
            
            if(maxHeightPixel < toTextHeightPixel){
                setHeightInPixels(toTextHeightPixel);
            }

            setFile(file, new Position(0 + padding, fromTextHeightPixel + padding, size - padding, fromTextHeightPixel + size - padding));

            setText(text);
        }

        return this;
    }

    /**
     * CellController 작업을 종료하고 SheetController Instance를 반환한다.
     * @return SheetController
     */
    public SheetController finishWorkcell(){
        return sheetController;
    }
}
