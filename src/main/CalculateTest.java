package main;

import com.sun.deploy.util.ArrayUtil;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartRenderingInfo;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.entity.StandardEntityCollection;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PiePlot3D;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.renderer.category.BarRenderer3D;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.util.ArrayUtilities;

import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

/**
 * Created by Администратор on 03.06.2017.
 */
public class CalculateTest {

    ArrayList<double[]> tests = new ArrayList();
    ArrayList<double[]> testsSelf = new ArrayList();
    double[] averegePoints = new double[4];
    double[][] totalAverege = new double[2][4];
    double[][] finalPoints = new double[2][4];
    double[][] finalPointsSelf = new double[2][4];
    String nameOf = "";

    ExelWork ex1 = null;
    CalculateTest(int num, String list [][], String listSelf [][], String name) throws IOException {
        nameOf = name;
        if(num == 1){
            calculateWorkerTest(list, name);
            calculateSelfTest(listSelf, name);
            totalAverege[0][0]=(finalPoints[0][0]*4+finalPointsSelf[0][0])/5;
            totalAverege[0][1]=(finalPoints[0][1]*4+finalPointsSelf[0][1])/5;
            totalAverege[0][2]=(finalPoints[0][2]*4+finalPointsSelf[0][2])/5;           // 80% of points is regular answers and 20% is "self answers" going in total of 100
            totalAverege[0][3]=(finalPoints[0][3]*4+finalPointsSelf[0][3])/5;
            //-----------------------------------------------------------
            totalAverege[1][0]=(finalPoints[1][0]*4+finalPointsSelf[1][0])/5;
            totalAverege[1][1]=(finalPoints[1][1]*4+finalPointsSelf[1][1])/5;
            totalAverege[1][2]=(finalPoints[1][2]*4+finalPointsSelf[1][2])/5;
            totalAverege[1][3]=(finalPoints[1][3]*4+finalPointsSelf[1][3])/5;
            System.out.println(totalAverege[0][0]);
            System.out.println(totalAverege[0][1]);
            System.out.println(totalAverege[0][2]);
            System.out.println(totalAverege[0][3]);

            System.out.println(finalPointsSelf[0][0]);
            System.out.println(finalPointsSelf[0][1]);
            System.out.println(finalPointsSelf[0][2]);
            System.out.println(finalPointsSelf[0][3]);

            System.out.println(finalPoints[0][0]);
            System.out.println(finalPoints[0][1]);
            System.out.println(finalPoints[0][2]);
            System.out.println(finalPoints[0][3]);
        }else{

        }
        //for(int i=0; i<tests.size(); i++){
            //System.out.print(" "+tests.get(i));
        //}

    }

    public void setPieChart(){
        DefaultPieDataset pieDataset = new DefaultPieDataset();
        pieDataset.setValue("Лідерство та стресостійкість", totalAverege[1][0]);
        pieDataset.setValue("Комунікація та робота з персоналом", totalAverege[1][1]);
        pieDataset.setValue("Робота з інформацією",  totalAverege[1][2]);
        pieDataset.setValue("Гнучкість та орієнтація на результат",  totalAverege[1][3]);
        JFreeChart pieChart = ChartFactory.createPieChart("Співвідношення набраних балів між розділами", pieDataset);
        //PiePlot3D P = (PiePlot3D)pieChart.getPlot();


        ChartRenderingInfo pieInfo = new ChartRenderingInfo(new StandardEntityCollection());
        File pieChartFile = new File("PieChart1.png");
        try {
            ChartUtilities.saveChartAsPNG(pieChartFile, pieChart, 700, 500);
        } catch (IOException e) {
            e.printStackTrace();
        }

    }


    public void setCategoryChart() throws IOException {
        DefaultCategoryDataset categoryDataset = new DefaultCategoryDataset();
        categoryDataset.setValue( totalAverege[1][0], "Оцінка", "Лідерство та стресостійкість");
        categoryDataset.setValue( totalAverege[1][1], "Оцінка", "Комунікація та робота з персоналом");
        categoryDataset.setValue( totalAverege[1][2], "Оцінка", "Робота з інформацією");
        categoryDataset.setValue( totalAverege[1][3], "Оцінка", "Гнучкість та орієнтація на результат");
        JFreeChart categoryChart = ChartFactory.createBarChart3D("Оцінка за розділами", "Назва розділу", "Оцінка", categoryDataset, PlotOrientation.VERTICAL, false, true, false);


        CategoryPlot b = (CategoryPlot)categoryChart.getPlot();
        BarRenderer3D barRenderer = (BarRenderer3D)b.getRenderer();
        barRenderer.setSeriesPaint(0, Color.GREEN);
        NumberAxis rangeAxis = (NumberAxis) b.getRangeAxis();
        rangeAxis.setRange(0,4);

        ChartRenderingInfo barInfo = new ChartRenderingInfo(new StandardEntityCollection());
        File barChartFile = new File("BarChart1.png");
        ChartUtilities.saveChartAsPNG(barChartFile, categoryChart, 700, 500);
    }

    public void setExelObject(ExelWork exel) throws IOException {
        ex1 = exel;
        //ex1.setExelSells(totalAverege, nameOf);
        setPieChart();
        setCategoryChart();
        ex1.setFinalExel(finalPoints,finalPointsSelf,totalAverege,nameOf);
    }

    public void calculateSelfTest(String listSelf[][], String name){
        for(int i=0; i<listSelf.length; i++){
            double score = 0;
            int questionCount = 0;
            int partCount = 0;
            System.out.println("Started loop!");
            if(listSelf[i][2].equals(name)){
                System.out.println("FoundSelfName!");
                testsSelf.add(new double[4]);
                for(String str: listSelf[i]){
                    score += getAnswerPoints(str);
                    if(score == 0){

                    }else{
                        questionCount++;
                        if(questionCount==3){
                            testsSelf.get(0)[partCount]=score;
                            System.out.println("AddedScore");
                            score = 0;
                            questionCount=0;
                            partCount++;
                        }
                    }
                }
                break;
            }

        }

        System.out.println(" ");
        System.out.println("SelfTests");
        System.out.println(testsSelf.get(0)[0]);
        System.out.println(testsSelf.get(0)[1]);
        System.out.println(testsSelf.get(0)[2]);
        System.out.println(testsSelf.get(0)[3]);
        System.out.println(" ");

        calculateFinalSelfPoints();
    }

    public void calculateWorkerTest(String list [][], String name){
        int countNames = 0;
        for(int i=0; i<list.length; i++){
            double score = 0;
            int questionCount = 0;
            int partCount = 0;
            if(list[i][2].equals(name)){
                tests.add(new double[4]);
                for(String str: list[i]){
                    score += getAnswerPoints(str);
                    if(score == 0){

                    }else{
                        questionCount++;
                        if(questionCount==5){
                            tests.get(countNames)[partCount]=score;
                            score = 0;
                            questionCount=0;
                            partCount++;
                        }
                    }
                }
                partCount = 0;
                countNames++;
            }
        }

        calculateAveregePoints();
        calculateFinalPoints();
    }

    public void calculateAveregePoints(){
        double part1 = 0;
        double part2 = 0;
        double part3 = 0;
        double part4 = 0;
        double countAmountOfPerson = 0;
            for(double[] person: tests){
                part1 += person[0];
                part2 += person[1];
                part3 += person[2];
                part4 += person[3];
                countAmountOfPerson++;
            }
        System.out.println(" ");
        System.out.println("p1");
        System.out.println(part1);
        System.out.println(part2);
        System.out.println(part3);
        System.out.println(part4);
        System.out.println(" ");
        averegePoints[0] = part1/countAmountOfPerson;
        averegePoints[1] = part2/countAmountOfPerson;
        averegePoints[2] = part3/countAmountOfPerson;
        averegePoints[3] = part4/countAmountOfPerson;
    }

    public void calculateFinalPoints(){
        double numberAfterPoint = 0;
        for(int i=0; i<4; i++){
                finalPoints[0][i]=averegePoints[i]*4;
                double n=averegePoints[i]*4;
                if(n<50){                                                   //switching from 100 points max to 4 points max
                    numberAfterPoint=(n/0.5)*0.02;                          // from 0 to 50 - 2 points
                    finalPoints[1][i]=numberAfterPoint;
                }else if((n<75)&&(n>=50)){                                  //  from 50 to 75 - 1 point( +2 for 0 to 50 range passed )
                    numberAfterPoint=((n-50)/0.25)*0.01;
                    finalPoints[1][i]=numberAfterPoint+2;
                }else if((n<95)&&(n>=75)){                                   // from 75 to 95 - 1 point ( +3 for previous ranges )
                    numberAfterPoint=((n-75)/0.2)*0.01;
                    finalPoints[1][i]=numberAfterPoint+3;
                }else if (n>=95){                                                // 95+ - max 4 points
                    finalPoints[1][i]=4;
                }
            numberAfterPoint = 0;
        }
        System.out.println(finalPoints[1][0]);
        System.out.println(finalPoints[1][1]);
        System.out.println(finalPoints[1][2]);
        System.out.println(finalPoints[1][3]);
    }

    public void calculateFinalSelfPoints(){
        double numberAfterPoint = 0;
        for(int i=0; i<4; i++){
            finalPointsSelf[0][i]=(((testsSelf.get(0)[i]*0.1)*100)/3)*2;
            double n=(((testsSelf.get(0)[i]*0.1)*100)/3)*2;                     //switching from 100 points max to 4 points max
            if(n<50){                                                           // from 0 to 50 - 2 points
                numberAfterPoint=(n/0.5)*0.02;
                finalPointsSelf[1][i]=numberAfterPoint;
            }else if((n<75)&&(n>=50)){                                              //  from 50 to 75 - 1 point( +2 for 0 to 50 range passed )
                numberAfterPoint=((n-50)/0.25)*0.01;
                finalPointsSelf[1][i]=numberAfterPoint+2;
            }else if((n<95)&&(n>=75)){                                              // from 75 to 95 - 1 point ( +3 for previous ranges )
                numberAfterPoint=((n-75)/0.2)*0.01;
                finalPointsSelf[1][i]=numberAfterPoint+3;
            }else if (n>=95){                                                       // 95+ - max 4 points
                finalPointsSelf[1][i]=4;
            }
            numberAfterPoint = 0;
        }
        System.out.println(finalPointsSelf[1][0]);
        System.out.println(finalPointsSelf[1][1]);
        System.out.println(finalPointsSelf[1][2]);
        System.out.println(finalPointsSelf[1][3]);
    }

    public double getAnswerPoints(String answer){
        double points = 0;
        try {
            if(answer.equals("ніколи")){
                points = 1.25;
            }else if(answer.equals("дуже рідко")){
                points = 2.50;
            }else if(answer.equals("часто")){
                points = 3.75;
            }else if(answer.equals("завжди")){
                points = 5;
            }
        }catch (NullPointerException e){

        }

        return points;
    }
}
