package excel;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelController {
    private Workbook workbook;
    private Font defaultFont;

    private Map<String, Integer> imageIndexMap = new HashMap<>(); // key: imageKey(사용자 지정), value: imageNumber(Workbook.addPicture())
    private Map<String, Integer> fileIndexMap = new HashMap<>(); // key: fileName(사용자 지정), value: fileNumber(Workbook.addOlePackage())
    private SheetController worksheetController;
    private List<SheetController> sheetControllerList = new ArrayList<>();

    private void destoryFields(){
        workbook = null;
        defaultFont = null;
    }

    /**
     * ExcelController의 생성자.
     * 내부적으로 Workbook과 Sheet를 생성한다.
     * @param sheetName 생성될 sheet의 이름
     */
    public ExcelController(){
        workbook = new XSSFWorkbook();
        addSheet();
        selectWorksheet(0);
        defaultFont = workbook.getFontAt(0);
        defaultFont.setFontName(Base.BASE_FONT_NAME);
        defaultFont.setFontHeightInPoints(Base.BASE_FONT_HEIGHT_POINTS);
        registBaseIconImage();
    }

    /**
     * workbook을 반환한다.
     * @return workbook
     */
    protected Workbook getWorkbook(){
        return workbook;
    }

    /**
     * imageIndexMap을 반환한다.
     * key: imageKey(사용자 지정), value: imageNumber(Workbook.addPicture())
     * @return imageIndexMap
     */
    protected Map<String, Integer> getImageIndexMap(){
        return imageIndexMap;
    }
    
    /**
     * fileIndexMap을 반환한다.
     * key: fileName(사용자 지정), value: fileNumber(Workbook.addOlePackage())
     * @return
     */
    protected Map<String, Integer> getFileIndexMap(){
        return fileIndexMap;
    }

    /**
     * 작업한 Workbook을 ByteArrayOutputStream으로 반환한다.
     * @return 작업한 Workbook을 ByteArrayOutputStream으로 반환한다.
     * @throws IOException
     */
    public ByteArrayOutputStream getByteArrayOutputStream() throws IOException{
        ByteArrayOutputStream result = new ByteArrayOutputStream();
        workbook.write(result);
        return result;
    }

    /**
     * Workbook을 close한다.
     * @throws IOException
     */
    public void close() throws IOException{
        workbook.close();
        destoryFields();
    }

    /**
     * 작업한 Workbook을 ByteArrayOutputStream으로 반환하고,
     * Workbook을 close한다.
     * @return 작업한 Workbook을 ByteArrayOutputStream으로 반환한다.
     * @throws IOException
     */
    public ByteArrayOutputStream getByteArrayOutputStreamAndClose() throws IOException{
        ByteArrayOutputStream result = getByteArrayOutputStream();
        close();
        return result;
    }

    /**
     * Sheet들의 이름을 순차적으로 반환한다.
     * @return Sheet들의 이름 목록들
     */
    public List<String> getSheetNameList(){
        return sheetControllerList.stream().map(s -> s.getSheetName()).collect(Collectors.toList());
    }

    /**
     * 현재 작업 중인 Sheet의 이름을 반환한다.
     * @return Sheet 이름
     */
    public String getWorksheetName(){
        return worksheetController.getSheetName();
    }

    /**
     * SheetController의 index를 반환한다.
     * SheetController의 없으면 -1을 반환한다.
     * @return SheetController의 index
     */
    public Integer getSheetControllerIndex(SheetController sheetController){
        int index = -1;
        for(int i = 0; i < sheetControllerList.size(); i++){
            if(sheetController == sheetControllerList.get(i)){
                index = i;
                break;
            }
        }
        return index;
    }

    /**
     * 작업중인 Sheet의 이름을 sheetName으로 설정한다.
     * @param sheetName 변경할 Sheet의 이름
     * @return this
     */
    public ExcelController setWorksheetName(final String sheetName){
        workbook.setSheetName(getSheetControllerIndex(worksheetController), sheetName);
        return this;
    }

    /**
     * Sheet를 추가한다.
     * @return this
     */
    public ExcelController addSheet(){
        SheetController sheetController = new SheetController(this);
        sheetControllerList.add(sheetController);
        return this;
    }

    /**
     * index번호로 작업중인 Sheet를 설정한다.
     * sheetIndex에 해당하는 Sheet가 존재하지 않을 경우, IlleaglArgumentException 예외를 발생한다.
     * @param sheetIndex Sheet 번호(0부터 시작)
     * @return SheetController
     */
    public SheetController selectWorksheet(final int sheetIndex) throws IllegalArgumentException{
        if(sheetIndex < 0 || sheetIndex >= sheetControllerList.size()){
            throw new IllegalArgumentException("sheetIndex에 해당하는 Sheet가 존재하지 않습니다.");
        }

        worksheetController = sheetControllerList.get(sheetIndex);

        return worksheetController;
    }

    /**
     * createObjectData에 쓰일 기본 아이콘 이미지를 등록한다.
     * icon 이미지 파일은 /static/poi에서 불러온다.
     */
    private void registBaseIconImage(){
        for(FileFormat fileFormat : FileFormat.values()){
            String iconFileName = fileFormat.getIconName();
            byte[] imageBytes = null;
            try {
                imageBytes = ClassLoader.getSystemClassLoader().getResourceAsStream("icon/" + iconFileName).readAllBytes();
            } catch (IOException e) {
                e.printStackTrace();
            }
            int imageIndex = workbook.addPicture(imageBytes, ImageFormat.PICTURE_TYPE_PNG.getValue());
            imageIndexMap.put(iconFileName, imageIndex);
        }
    }

    /**
     * Image 처리를 위해 사용되는 등록된 ImageKey들을 Set형태로 리턴한다.
     * @return Image 처리를 위해 사용되는 등록된 ImageKey들을 Set형태로 리턴한다.
     */
    public Set<String> getImageKeySet(){
        return imageIndexMap.keySet();
    }
}