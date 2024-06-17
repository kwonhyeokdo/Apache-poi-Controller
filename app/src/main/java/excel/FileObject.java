package excel;

public class FileObject {
    private byte[] fileByteArray;
    private FileFormat fileFormat;
    private String fileName;
    
    public FileObject(
        byte[] fileByteArray,
        FileFormat fileFormatEnum,
        String fileName
    ) {
        this.fileByteArray = fileByteArray;
        this.fileFormat = fileFormatEnum;
        this.fileName = fileName;
    }

    /**
     * Getter
     * @return embeddedFileByteArray
     */
    public byte[] getFileByteArray() {
        return fileByteArray;
    }

    /**
     * Getter
     * @return embeddedFileByteArray
     */
    public FileFormat getFileFormat() {
        return fileFormat;
    }

    /**
     * Getter
     * @return embeddedFileByteArray
     */
    public String getFileName() {
        return fileName;
    }
}
