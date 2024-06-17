package excel;

import java.util.Map;

public class Base {
    protected static final String BASE_FONT_NAME = "맑은 고딕";
    protected static final short BASE_FONT_HEIGHT_POINTS = 10;
    protected static final int POI_WIDTH_UNIT = 256;
    private static final Map<Integer, Integer> CHARACTER_WIDTH_PIXELS_MAP = Map.ofEntries(
        Map.entry(5, 4),
        Map.entry(6, 4),
        Map.entry(7, 5),
        Map.entry(8, 6),
        Map.entry(9, 7),
        Map.entry(10, 7),
        Map.entry(11, 8),
        Map.entry(12, 9),
        Map.entry(13, 9),
        Map.entry(14, 10),
        Map.entry(15, 11),
        Map.entry(16, 12),
        Map.entry(17, 13),
        Map.entry(18, 13),
        Map.entry(19, 14),
        Map.entry(20, 15),
        Map.entry(21, 15)
        // Map.entry(22, 16.0),
        // Map.entry(23, 17.0),
        // Map.entry(24, 18.0)
    );
    private static Map<Integer, Integer> CHARACTER_HEIGHT_PIXELS_MAP = Map.ofEntries(
        Map.entry(5, 12),
        Map.entry(6, 13),
        Map.entry(7, 13),
        Map.entry(8, 16),
        Map.entry(9, 16),
        Map.entry(10, 18),
        Map.entry(11, 22),
        Map.entry(12, 23),
        Map.entry(13, 26),
        Map.entry(14, 27),
        Map.entry(15, 32),
        Map.entry(16, 35),
        Map.entry(17, 35),
        Map.entry(18, 35),
        Map.entry(19, 40),
        Map.entry(20, 42),
        Map.entry(21, 42)
        // Map.entry(22, 45),
        // Map.entry(23, 47),
        // Map.entry(24, 51)
    );

    protected static int getCharacterWidthPixels(double points){
        return CHARACTER_WIDTH_PIXELS_MAP.get((int)points);
    }

    protected static int getCharacterHeightPixels(double points){
        return CHARACTER_HEIGHT_PIXELS_MAP.get((int)points);
    }
}
