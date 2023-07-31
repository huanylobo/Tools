package org.example;
import java.awt.*;
import java.awt.event.InputEvent;
import java.lang.Object;
/**
 * @author hhhb
 * @date 2023/7/31
 */
public class MouseTools {
    public static void main(String[] args) throws Exception{
        Thread.sleep(3000);
        double[] position = getPosition();
        tool(position[0],position[1]);
    }
    public static void tool(double positionX,double positionY) throws AWTException {
        try {
            Robot robot = new Robot();
            robot.mouseMove((int)positionX,(int)positionY);
            robot.mousePress(InputEvent.BUTTON1_MASK);
            robot.mouseRelease(InputEvent.BUTTON1_MASK);
        } catch (AWTException e) {
            e.printStackTrace();
        } finally {
            System.out.println("dianjijieshu");
        }
    }

    public static double[] getPosition(){
        double[] position=new double[2];
        Point p  = MouseInfo.getPointerInfo().getLocation();
        position[0]=p.getX();
        position[1]=p.getY();
        System.out.println("X:"+p.getX() + "\n---\nY:" +p.getY());
        return position;
    }
}
