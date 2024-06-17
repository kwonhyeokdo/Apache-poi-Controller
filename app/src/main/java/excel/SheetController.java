package excel;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class SheetController {
    private final ExcelController excelController;
    private final Workbook workbook;
    private final Sheet worksheet;

    private final Map<String, CellController> cellControllerMap = new HashMap<>();
    private CellController workcellController;

    protected SheetController(ExcelController excelController){
        this.excelController = excelController;
        workbook = excelController.getWorkbook();
        worksheet = workbook.createSheet();
    }

    /**
     * worksheet을 반환한다.
     * @return worksheet
     */
    protected Sheet getWorksheet(){
        return worksheet;
    }

    protected Drawing<?> getWorkdrawing(){
        Drawing<?> drawing = worksheet.getDrawingPatriarch();
        if(drawing == null){
            drawing = worksheet.createDrawingPatriarch();
        }
        return drawing;
    }

    /**
     * ExcelController Instance를 반환한다.
     * @return ExcelController
     */
    protected ExcelController getExcelController(){
        return excelController;
    }

    /**
     * workSheet의 SheetName을 반환한다.
     * @return SheetName
     */
    public String getSheetName(){
        return worksheet.getSheetName();
    }

    public SheetController setSheetName(final String sheetName){
        int sheetControllerIndex = excelController.getSheetControllerIndex(this);
        workbook.setSheetName(sheetControllerIndex, sheetName);

        return this;
    }

    /**
     * Excel의 기본 Column Width를 columnWidth 값으로 설정한다.
     * Apache-Poi의 Sheet의 setDefaultColumnWidth() 특성상 Width 결과가 columnWidth 값으로 정확히 적용되지 않을 수 있다.
     * @param columnWidth Excel의 Column Width 설정 단위랑 같다.
     * @return this
     */
    public SheetController setDefaultColumnWidth(final int columnWidth){
        worksheet.setDefaultColumnWidth(columnWidth);
        return this;
    }

    /**
     * Excel의 기본 Column Width를 pixels 값으로 설정한다.
     * @param pixels
     * @return this
     */
    public SheetController setDefaultColumnWidthInPixels(final int pixels){
        int width = UnitConverter.pixelsToPoiColumnWidth(pixels) / Base.POI_WIDTH_UNIT;
        setDefaultColumnWidth(width);
        return this;
    }

    /**
     * Excel의 기본 Row Height를 points 단위의 값으로 설정한다.
     * @param points Excel의 Row Height 설정 단위랑 같다.
     * @return this
     */
    public SheetController setDefaultRowHeightInPoints(final double points){
        worksheet.setDefaultRowHeightInPoints((float)points);
        return this;
    }

    /**
     * Excel의 기본 Row Height를 pixels 값으로 설정한다.
     * @param pixels
     * @return this
     */
    public SheetController setDefaultRowHeightInPixels(final int pixels){
        double point = UnitConverter.pixelsToPoints(pixels);
        setDefaultRowHeightInPoints(point);
        return this;
    }

    /**
     * columnIndex번째 Column의 Width를 columnWidth으로 변경한다.
     * @param columnIndex 변경할 Column의 번호(0부터 시작).
     * @param columnWidth Excel의 Column Width 설정 단위랑 같다.
     * @return this
     */
    public SheetController setColumnWidth(final int columnIndex, final double columnWidth){
        int poiColumnWidth = UnitConverter.columnWidthToPoiColumnWidth(columnWidth);
        worksheet.setColumnWidth(columnIndex, poiColumnWidth);
        return this;
    }

    /**
     * columnIndex번째 Column의 Width를 pixels으로 변경한다.
     * @param columnIndex 변경할 Column의 번호(0부터 시작).
     * @param pixel
     * @return this
     */
    public SheetController setColumnWidthInPixels(final int columnIndex, final int pixels){
        int poiColumnWidth = UnitConverter.pixelsToPoiColumnWidth(pixels);
        worksheet.setColumnWidth(columnIndex, poiColumnWidth);
        return this;
    }

    /**
     * rowIndex에 해당하는 Row의 Height를 points으로 변경한다.
     * @param rowIndex 변경할 Row의 번호(0부터 시작).
     * @param points Excel의 Row Height 설정 단위랑 같다.
     * @return this
     */
    public SheetController setRowHeightInPoints(final int rowIndex, final int points){
        Row row = getRow(rowIndex);
        row.setHeight(UnitConverter.pointsToPoiHeight(points));
        return this;
    }

    /**
     * rowIndex에 해당하는 Row의 Height를 pixels으로 변경한다.
     * @param rowIndex 변경할 Row의 번호(0부터 시작).
     * @param pixel
     * @return this
     */
    public SheetController setRowHeightInPixels(final int rowIndex, final int pixels){
        Row row = getRow(rowIndex);
        row.setHeight(UnitConverter.pixelsToPoiHeight(pixels));
        return this;
    }

    /**
     * rowIndex에 해당하는 Row를 반환한다.
     * @param rowIndex
     * @return rowIndex에 해당하는 Row
     */
    protected Row getRow(final int rowIndex){
        Row row = worksheet.getRow(rowIndex);
        if(row == null){
            row = worksheet.createRow(rowIndex);
        }
        return row;
    }

    /**
     * cellControllerMap에 쓰이는 Key를 생성한다.
     * @param rowIndex Row의 번호(0부터 시작).
     * @param colInex Column의 번호(0부터 시작).
     * @return "R{rowIndex}C{colIndex}"
     */
    private String getCellControllerKey(final int rowIndex, final int colInex){
        return "R" + rowIndex + "C" + colInex;
    }

    /**
     * 작업할 Cell을 선택한다.
     * @param rowIndex Row의 번호(0부터 시작).
     * @param colIndex Column의 번호(0부터 시작).
     * @return CellController
     */
    public CellController selectCell(final int rowIndex, final int colIndex){
        String cellControllerKey = getCellControllerKey(rowIndex, colIndex);

        if(cellControllerMap.containsKey(cellControllerKey)){
            return cellControllerMap.get(cellControllerKey);
        }else{
            workcellController = new CellController(this, rowIndex, colIndex);
            cellControllerMap.put(cellControllerKey, workcellController);
            return workcellController;
        }
    }
    

    /**
     * Cell을 Merge한다.
     * @param startRowIndex 시작 Row Index(0부터 시작)
     * @param endRowIndex 종료 Row Index(0부터 시작)
     * @param startColIndex 시작 Col Index(0부터 시작)
     * @param endColInex 종료 Col Index(0부터 시작)
     * @return this
     */
    public SheetController mergedRegion(int startRowIndex, int endRowIndex, int startColIndex, int endColInex){
        worksheet.addMergedRegion(new CellRangeAddress(startRowIndex, endRowIndex, startColIndex, endColInex));
        return this;
    }

    /**
     * Cell을 Merge한다.
     * @param startRowIndex 시작 Row Index(0부터 시작)
     * @param endRowIndex 종료 Row Index(0부터 시작)
     * @param startColIndex 시작 Col Index(0부터 시작)
     * @param endColInex 종료 Col Index(0부터 시작)
     * @return Mearge된 영역의 Cell을 조정할 수 있는 CellController 인스턴스를 반환한다.
     */
    public CellController mergedRegionAndSelectCell(int startRowIndex, int endRowIndex, int startColIndex, int endColInex){
        worksheet.addMergedRegion(new CellRangeAddress(startRowIndex, endRowIndex, startColIndex, endColInex));
        return selectCell(startRowIndex, startColIndex);
    }

    /**
     * Cell을 Merge한다.
     * @param startCell ex) "A1"
     * @param endCell ex) "C1"
     * @return this
     */
    public SheetController mergedRegion(String startCell, String endCell){
        worksheet.addMergedRegion(CellRangeAddress.valueOf(startCell + ":" + endCell));
        return this;
    }

    /**
     * SheetController 작업을 종료하고 ExcelController Instance를 반환한다.
     * @return ExcelController
     */
    public ExcelController finishWorksheet(){
        return excelController;
    }
}
