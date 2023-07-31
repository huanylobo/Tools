package org.example;

import java.util.concurrent.CompletableFuture;

/**
 * 线程学习
 * @author test
 * @date 2023/6/28
 */
public class Test extends Thread{
    @Override
    public void run() {
        for (int i = 0; i < 20; i++) {
            System.out.println("我在看代码"+i+ " " + Thread.currentThread().getName());
        }
    }

    public static void main(String[] args) {
        //main线程  主线程
        //创建一个线程对象
        System.out.println(Thread.currentThread().getName());
        Test test1 = new Test();
        test1.setName("线程1");
        Test test2 = new Test();
        test2.setName("线程2");
        //调用start开启线程
        test1.start();
        test2.start();
        System.out.println(Thread.currentThread().getName());
    }

}
