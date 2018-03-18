package main;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;

/**
 * Created by Администратор on 03.06.2017.
 */
public class GUI extends JFrame {

    ExelWork exelrun = null;


    JFrame mainFrame = new JFrame();
    JPanel mainMenu = new JPanel();
    JLabel background = null;
    JButton buttonDialog = new JButton("Укажіть шлях до .xlsx файлів");
    JButton buttonNext = new JButton("Менеджер");
    JButton buttonNext2 = new JButton("Ст Продавець");
    JButton buttonFinish = new JButton("Створити .xlsx файл");

    String exelPath = "";


    JFileChooser pathSelect = new JFileChooser();
    JScrollPane scrollNames = null;
    JList listOfNames = null;


    ImageIcon backgroundPic= new ImageIcon("pic\\background.png");

    String job = null;

    GUI(){
        mainFrame.setTitle("Welcome");
        mainFrame.setName("mainFrame");
        mainFrame.setBounds(400, 150, 600, 400);
        mainFrame.setLayout(null);
        mainFrame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        mainFrame.setBackground(Color.black);
        mainFrame.setResizable(false);
        mainFrame.setVisible(true);
        //mainMenu.setBackground(Color.BLACK);
        setMainMenu();
        mainFrame.setContentPane(mainMenu);
        initListener();
    }

    private void initListener(){
        buttonDialog.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                pathSelect.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
                pathSelect.showDialog(mainMenu, "Open");
                buttonNext.setVisible(true);
                buttonNext2.setVisible(true);
            }
        });
        buttonNext.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                job = "Менеджер";
                buttonDialog.setVisible(false);
                exelPath = pathSelect.getCurrentDirectory().getPath();
                System.out.println(exelPath);
                try {
                    exelrun = new ExelWork(exelPath,job);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                buttonNext.setVisible(false);
                buttonNext2.setVisible(false);

                scrollNames = new JScrollPane();
                listOfNames = new JList(exelrun.getMainList());
                scrollNames.setBounds(150,145,300,100);
                listOfNames.setBounds(150,145,300,100);
                listOfNames.setVisibleRowCount(6);
                background.add(listOfNames);
                background.add(scrollNames);
                scrollNames.setViewportView(listOfNames);
                listOfNames.setVisible(false);
                listOfNames.setVisible(true);
                scrollNames.setVisible(true);
                buttonFinish.setVisible(true);
            }
        });
        buttonNext2.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                job = "Ст Продавець";
                buttonDialog.setVisible(false);
                exelPath = pathSelect.getCurrentDirectory().getPath();
                System.out.println(exelPath);
                try {
                    exelrun = new ExelWork(exelPath, job);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                buttonNext.setVisible(false);
                buttonNext2.setVisible(false);

                scrollNames = new JScrollPane();
                listOfNames = new JList(exelrun.getMainList());
                scrollNames.setBounds(150,145,300,100);
                listOfNames.setBounds(150,145,300,100);
                listOfNames.setVisibleRowCount(6);
                background.add(listOfNames);
                background.add(scrollNames);
                scrollNames.setViewportView(listOfNames);
                listOfNames.setVisible(false);
                listOfNames.setVisible(true);
                scrollNames.setVisible(true);
                buttonFinish.setVisible(true);
            }
        });
        buttonFinish.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String finalName = listOfNames.getSelectedValue().toString();
                CalculateTest calc = null;
                try {
                    calc = new CalculateTest(1, exelrun.getStandartList(), exelrun.getSelfList(), finalName);
                    System.out.println("Calculations finished!");
                    calc.setExelObject(exelrun);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }

            }
        });
    }

    public void setMainMenu(){
        mainMenu.setLayout(null);
        background = new JLabel(backgroundPic);
        background.setBounds(0, 0, 600, 400);
        background.setLayout(null);
        background.setBackground(Color.black);
        mainMenu.add(background);

        buttonNext.setBounds(130,300,130,30);
        buttonNext.setVisible(false);
        buttonNext2.setBounds(330,300,130,30);
        buttonNext2.setVisible(false);
        buttonDialog.setBounds(190,200,220,50);
        //buttonDialog.setBackground(Color.BLACK);
        buttonDialog.setVisible(true);

        buttonFinish.setBounds(200,280,200,35);
        buttonFinish.setVisible(false);

        background.add(buttonDialog);
        background.add(buttonNext);
        background.add(buttonNext2);
        background.add(buttonFinish);



    }



}
