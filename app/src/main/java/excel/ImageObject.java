package excel;

import java.io.ByteArrayInputStream;
import java.io.IOException;

import javax.imageio.ImageIO;
import java.awt.Image;

public class ImageObject {
    private byte[] imageByteArray;
    private ImageFormat imageFormat;
    private String imageKey;
    private Image image;
    
    public ImageObject(byte[] imageByteArray, ImageFormat imageFormat, String imageKey) throws IOException{
        this.imageByteArray = imageByteArray;
        this.imageFormat = imageFormat;
        this.imageKey = imageKey;
        this.image = ImageIO.read(new ByteArrayInputStream(imageByteArray));
    }

    /**
     * Getter
     * @return imageByteArray
     */
    public byte[] getImageByteArray() {
        return imageByteArray;
    }

    /**
     * Getter
     * @return imageFormat
     */
    public ImageFormat getImageFormat() {
        return imageFormat;
    }

    /**
     * Getter
     * @return imageKey
     */
    public String getImageKey() {
        return imageKey;
    }

    /**
     * Image의 Width를 pixels로 반환한다.
     * @return image의 width(Pixels)
     */
    public int getWidth(){
        return image.getWidth(null);
    }

    /**
     * Image의 Height를 pixels로 반환한다.
     * @return image의 height(pixels)
     */
    public int getHeight(){
        return image.getHeight(null);
    }
}
