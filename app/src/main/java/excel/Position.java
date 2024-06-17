package excel;

public class Position {
    private int dx1;
    private int dy1;
    private int dx2;
    private int dy2;

    public Position(int dx1, int dy1, int dx2, int dy2) {
        if(dx1 < 0 || dy1 < 0 || dx2 < 0 || dy2 < 0){
            throw new IllegalArgumentException("Values of the parameter must be greater than or equal to zero.");
        }
        this.dx1 = dx1;
        this.dy1 = dy1;
        this.dx2 = dx2;
        this.dy2 = dy2;
    }

    public int getDx1() {
        return dx1;
    }

    public void setDx1(int dx1) {
        this.dx1 = dx1;
    }

    public int getDy1() {
        return dy1;
    }

    public void setDy1(int dy1) {
        this.dy1 = dy1;
    }

    public int getDx2() {
        return dx2;
    }

    public void setDx2(int dx2) {
        this.dx2 = dx2;
    }

    public int getDy2() {
        return dy2;
    }

    public void setDy2(int dy2) {
        this.dy2 = dy2;
    }
}
