package excel;

/**
 * 96 DPI 기준
 */
public class UnitConverter {
    private static final double POINTS_PER_PIXELS = 0.75;
    private static final int POI_COLUMN_WIDTH_UNIT = 256;
    private static final short POI_ROW_HEIGHT_UNIT = 20;
    
    /**
     * pixels를 Points로 변환한다.
     * @param pixels
     * @return 변환된 Points
     */
    protected static double pixelsToPoints(int pixels){
        return POINTS_PER_PIXELS * pixels;
    }


    /**
     * points를 Pixels로 변환한다.
     * @param point
     * @return 변환된 Pixels
     */
    protected static int pointsToPixels(double points){
        return (int)(points / POINTS_PER_PIXELS);
    }

    /**
     * pixels을 Excel의 ColumnWidth으로 변경한다.
     * @param pixels
     * @return ColumnWidth
     */
    protected static int pixelsToPoiColumnWidth(final int pixels){
        int columnWidth = (int)((double)pixels / Base.getCharacterWidthPixels(Base.BASE_FONT_HEIGHT_POINTS) * POI_COLUMN_WIDTH_UNIT);
        return columnWidth;
    }

    /**
     * columnWidth를 PoiColumnWidth으로 변경한다.
     * PaddingColumnWidth = (columnWidth * characterWidthPixels + 5.0) / characterWidthPixels
     * 참고: https://stackoverflow.com/questions/70121858/can-not-accurately-set-excel-column-width-apache-poi
     * PoiColumnWidth = PaddingColumnWidth * 256
     * @param columnWidth
     * @return PoiColumnWidth
     */
    protected static int columnWidthToPoiColumnWidth(final double columnWidth){
        // int paddingColumnWidth = columnWidthToPaddingColumnWidth(columnWidth);
        // return paddingColumnWidth * POI_COLUMN_WIDTH_UNIT;
        int characterWidthPixels = Base.getCharacterWidthPixels(Base.BASE_FONT_HEIGHT_POINTS);
        int poiColumnWidth = (int)Math.round(
            (double)(columnWidth * characterWidthPixels + 5.0) / characterWidthPixels * POI_COLUMN_WIDTH_UNIT
        );
        return poiColumnWidth;
    }

    /**
     * points를 PoiHeight으로 변경한다.
     * PoiHeight = Height * 20
     * @param points
     * @return PoiHeight
     */
    public static short pointsToPoiHeight(int points) {
        return (short)(points * POI_ROW_HEIGHT_UNIT);
    }

    /**
     * pixels를 PoiHeight으로 변경한다.
     * PoiHeight = Height * 20
     * @param pixels
     * @return PoiHeight
     */
    public static short pixelsToPoiHeight(int pixels) {
        int points = (int)pixelsToPoints(pixels);
        return pointsToPoiHeight(points);
    }

    /**
     * Excel의 Column Width를 pixels로 변환한다.
     * @param columnWidth
     * @return Excel의 Column Width를 pixel로 변환한다.
     */
    static int columnWidthToPixels(int columnWidth){
        return (int)(columnWidth * Base.getCharacterWidthPixels(Base.BASE_FONT_HEIGHT_POINTS) / Base.POI_WIDTH_UNIT);
    }
}
