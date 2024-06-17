package excel;

public enum ImageFormat {
    PICTURE_TYPE_EMF(2), // Extended windows meta file
    PICTURE_TYPE_WMF(3), // Windows Meta File
    PICTURE_TYPE_PICT(4), // Mac PICT format
    PICTURE_TYPE_JPEG(5), // JPEG format
    PICTURE_TYPE_PNG(6), // PNG format
    PICTURE_TYPE_DIB(7); // Device independent bitmap

    private final int value;
    
    ImageFormat(int value){
        this.value = value;
    }

    public int getValue(){
        return this.value;
    }
}
