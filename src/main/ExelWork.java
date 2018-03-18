package main;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sun.nio.ch.IOUtil;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;

/**
 * Created by Администратор on 03.06.2017.
 */
public class ExelWork {

    String list[][];
    String listSelf[][];

    Workbook wb = null;
    Workbook wbSelf = null;
    Workbook wb1 =  null;
    Sheet sheetMain1 = null;
    Sheet sheetMain2 = null;
    Sheet sheetMain3 = null;
    Sheet sheet1 = null;
    Row row1 = null;
    Row row2 = null;
    Row rowSet = null;
    Cell cellSet = null;
    Row rowSet2 = null;
    Cell cellSet2 = null;
    Cell cell1 = null;
    Cell cell2 = null;
    FileOutputStream fos = null;
    String exelWorkPath = "";
    String job1 = null;
    DecimalFormat doubleFormat = new DecimalFormat("#.###");

    FileInputStream pictureStream1 = null;
    FileInputStream pictureStream2 = null;

    Workbook finalwb =  null;

    ExelWork(String path, String job) throws IOException {
        job1 = job;
        exelWorkPath = path;

        if(job.equals("Менеджер")){
            wb = new XSSFWorkbook(new FileInputStream(path+"\\Опитувальний бланк керівника (Відповіді).xlsx"));
            wbSelf = new XSSFWorkbook(new FileInputStream(path+"\\Самооцінювання (Відповіді).xlsx"));
        }else if(job.equals("Ст Продавець")){
            wb = new XSSFWorkbook(new FileInputStream(path+"\\ст.пр. опитувальний бланк  (Відповіді).xlsx"));
            wbSelf = new XSSFWorkbook(new FileInputStream(path+"\\Самооцінювання Ст.продавець (Відповіді).xlsx"));
        }else{
            System.out.println("Path or some shit didn't work m8");
        }
        System.out.println(job);
        Sheet sheet = wb.getSheetAt(0);
        Sheet sheetSelf = wbSelf.getSheetAt(0);
        Row row0 = sheet.getRow(0);
        Row rowSelf0 = sheetSelf.getRow(0);

        wb1 = new XSSFWorkbook();
        sheet1 = wb1.createSheet("first");

        int RowCount=0;
        int CellCount=0;
        int number = 1;
        int i;
        list = new String[sheet.getLastRowNum()+1][row0.getLastCellNum()+1];
        System.out.println("length: "+row0.getLastCellNum());
        System.out.println("lengthRow: "+sheet.getLastRowNum());
        listSelf = new String[sheetSelf.getLastRowNum()+1][rowSelf0.getLastCellNum()+1];
        //for( i=sheet.getFirstRowNum(); i<=sheet.getLastRowNum(); i++ ){
        for(Row row: wb.getSheetAt(0)){                                                     //reading answers from input exel file
            for(Cell cell : row){
                list[RowCount][CellCount]=getCellType(cell);
                CellCount++;
            }
            CellCount=0;
            RowCount++;
        }
        RowCount = 0;
        CellCount=0;
        for(Row rowSelf: wbSelf.getSheetAt(0)){                                                     //reading "self answers" from input exel file
            for(Cell cellSelf : rowSelf){
                listSelf[RowCount][CellCount]=getCellType(cellSelf);
                CellCount++;
            }
            CellCount=0;
            RowCount++;
        }
        //System.out.println(list[10][2]);
        //System.out.println(list.length);

    }

    public void setFinalExel(double[][] tests, double[][] testsSelf, double[][] averegeTests, String name ) throws IOException {
        finalwb = new XSSFWorkbook(new FileInputStream(exelWorkPath+"\\Бланк обробки даних SS17.xlsx"));
        sheetMain1 = finalwb.getSheetAt(0);
        sheetMain2 = finalwb.getSheetAt(1);
        sheetMain3 = finalwb.getSheetAt(2);


        //------------------------First page records------------------------------1!-----------------
        setCellValueString(3,4,name);

        setCellValueString(4,4,job1);

        for(int i=1; i<listSelf.length; i++){
            if(listSelf[i][2].equals(name)){
                setCellValueString(5,4,listSelf[i][1]);
            }
        }

        setCellValueDouble(8, 5, averegeTests[1][0]);
        setCellValuePersent(8, 6, doubleFormat.format(averegeTests[0][0]));

        setCellValueDouble(9, 5, averegeTests[1][1]);
        setCellValuePersent(9, 6, doubleFormat.format(averegeTests[0][1]));

        setCellValueDouble(10, 5, averegeTests[1][2]);
        setCellValuePersent(10, 6, doubleFormat.format(averegeTests[0][2]));

        setCellValueDouble(11, 5, averegeTests[1][3]);
        setCellValuePersent(11, 6, doubleFormat.format(averegeTests[0][3]));

        setCellValueDouble(12, 7, (averegeTests[1][0]+averegeTests[1][1]+averegeTests[1][2]+averegeTests[1][3])/4);

        //------------------------Second page records------------------------------2!-----------------

        //-----------------images
        pictureStream1 = new FileInputStream(exelWorkPath+"\\BarChart1.png");
        pictureStream2 = new FileInputStream(exelWorkPath+"\\PieChart1.png");
        CreationHelper helper = finalwb.getCreationHelper();
        Drawing drawing = sheetMain2.createDrawingPatriarch();

        ClientAnchor anchor1 = helper.createClientAnchor();
        anchor1.setAnchorType(ClientAnchor.MOVE_AND_RESIZE);

        ClientAnchor anchor2 = helper.createClientAnchor();
        anchor2.setAnchorType(ClientAnchor.MOVE_AND_RESIZE);

        int pictureIndex1 = finalwb.addPicture(IOUtils.toByteArray(pictureStream1),finalwb.PICTURE_TYPE_PNG);
        int pictureIndex2 = finalwb.addPicture(IOUtils.toByteArray(pictureStream2),finalwb.PICTURE_TYPE_PNG);

        anchor1.setCol1(0);
        anchor1.setRow1(0);
        anchor1.setCol1(0);
        anchor1.setRow2(0);

        Picture pic1 = drawing.createPicture(anchor1, pictureIndex1);
        pic1.resize();

        anchor2.setCol1(7);
        anchor2.setRow1(0);
        anchor2.setCol1(7);
        anchor2.setRow2(0);

        Picture pic2 = drawing.createPicture(anchor2, pictureIndex2);
        pic2.resize();

        //----------general
        setCellValueDoubleSecondSheet(29, 3, tests[1][0]);
        setCellValuePersentSecondSheet(29, 4, doubleFormat.format(tests[0][0]));

        setCellValueDoubleSecondSheet(30, 3, tests[1][1]);
        setCellValuePersentSecondSheet(30, 4, doubleFormat.format(tests[0][1]));

        setCellValueDoubleSecondSheet(31, 3, tests[1][2]);
        setCellValuePersentSecondSheet(31, 4, doubleFormat.format(tests[0][2]));

        setCellValueDoubleSecondSheet(32, 3, tests[1][3]);
        setCellValuePersentSecondSheet(32, 4, doubleFormat.format(tests[0][3]));

        //----------own
        setCellValueDoubleSecondSheet(29, 5, testsSelf[1][0]);
        setCellValuePersentSecondSheet(29, 6, doubleFormat.format(testsSelf[0][0]));

        setCellValueDoubleSecondSheet(30, 5, testsSelf[1][1]);
        setCellValuePersentSecondSheet(30, 6, doubleFormat.format(testsSelf[0][1]));

        setCellValueDoubleSecondSheet(31, 5, testsSelf[1][2]);
        setCellValuePersentSecondSheet(31, 6, doubleFormat.format(testsSelf[0][2]));

        setCellValueDoubleSecondSheet(32, 5, testsSelf[1][3]);
        setCellValuePersentSecondSheet(32, 6, doubleFormat.format(testsSelf[0][3]));

        //----------averege
        setCellValueDoubleSecondSheet(29, 9, averegeTests[1][0]);
        setCellValuePersentSecondSheet(29, 10, doubleFormat.format(averegeTests[0][0]));

        setCellValueDoubleSecondSheet(30, 9, averegeTests[1][1]);
        setCellValuePersentSecondSheet(30, 10, doubleFormat.format(averegeTests[0][1]));

        setCellValueDoubleSecondSheet(31, 9, averegeTests[1][2]);
        setCellValuePersentSecondSheet(31, 10, doubleFormat.format(averegeTests[0][2]));

        setCellValueDoubleSecondSheet(32, 9, averegeTests[1][3]);
        setCellValuePersentSecondSheet(32, 10, doubleFormat.format(averegeTests[0][3]));

        //----------Strings!-------------------------------

        //--------------------1
        if(averegeTests[0][0]<50){
            setCellValueStringSecondSheet(29,7,"" +
                    "По своїй натурі ви більш схильні приймати роль виконавця або підлеглого. Зазвичай ви невпевнені у своїх знаннях у тій чи іншій сферах, що значно уповільнює процес прийняття рішень та впливає на їх ефективність. Вам важко знаходити контакт з підлеглими та досягти порозуміння у складних ситуаціях. Спробуйте почати приймати рішення в тих областях, де відмова або невдача не стане критичною для вашої впевненості у собі і у своїх силах. Радимо активно шукати зворотній зв`язок з підлеглими, це допоможе зрозуміти наскільки ваші уявлення про мету та роль в колективі розходиться з уявленням ваших колег.");
        }else if((averegeTests[0][0]<70)&&(averegeTests[0][0]>=50)){
            setCellValueStringSecondSheet(29,7,"" +
                    "Вам важко діяти оганізовано та цілеспрямовано в умовах невизначенності. Відсутність порозуміння з колегами заважає безпроблемному вирішенні складних та конфліктних ситуацій. Сформовані вами цілі часто не мають чіткої мети, що заважає досягнути високих результатів. При формуванні мотивації радимо  переконатись у необхідності та, в першу чергу для самого себе, зрозуміти як робота кожного може якісно вплинути на досягнення спільної мети. Щоб досягти успіху у вирішенні конфліктних ситуацій спробуйте детальніше проаналізувати суть конфлікту та виявити шляши їх вирішення.");
        }else if((averegeTests[0][0]<90)&&(averegeTests[0][0]>=70)){
            setCellValueStringSecondSheet(29,7,"" +
                    "Ви безумовно є лідером, якого усі сприймають як частину команди, який направляє колектив для досягнення мети. Ви користуєтесь довірою своїх співробітників і чесність в цьому - головна ваша перевага. У стесових ситуаціях ви не втрачаєте сомоконтроль, влучно визначаєте методи вирішення будь-якої ситуації. Для удосконалення своїх якостей керівника радимо більшу увагу приділяти зворотньому зв`язку від персоналу.");
        }else if (averegeTests[0][0]>=90){
            setCellValueStringSecondSheet(29,7,"" +
                    "Ви є головним джерелом натхнення і мотивації для усього колективу. Ви безумовно є керівником, якого поважають, а не бояться. Вашому стилю керівництва властива завзятість, що є головним секретом досягнення бажаного результату. Ви завжди дієте організовано та цілеспрямовано в умовах невизначеності. Вам властивий високий рівень самоконтролю та сомокритики. Ви безперечно є тим лідером, на якого ваш персонал дивиться в пошуках впевненості та підтримки у складних ситуаціях.");
        }

        //------------------------2
        if(averegeTests[0][1]<50){
            setCellValueStringSecondSheet(30,7,"" +
                    "Неформальний лідер. Ви нерідко стаєте ініціатором конфліктних ситуацій і більш сприяєте їх подальшому розвитку, аніж вирішенню. Не вважаєте за потрібне делегувати повноваження. Зазвичай ставите свої інтереси вище за інтереси команди/компанії. Комунікація з колективом знаходиться на низькому рівні та потребує розвитку данної компетенції. Радимо переглянути стиль управління персоналом та звернути увагу на ефективність якісного делегування. Радимо приймати активнішу участь у вирішенні колективних питань та у робочих процесах загалом.");
        }else if((averegeTests[0][1]<70)&&(averegeTests[0][1]>=50)){
            setCellValueStringSecondSheet(30,7,"" +
                    "Зазвичай ви приймаєте участь в командній роботі виключно у присутності безпосереднього керівника. Співробітникам зазвичай важко знайти з вами спільну мову. Часто використовуєте робочий час для вирішення особистих питань. Вам характерна незацікавленість у досягненні цілей, тому часто ви перекладаєте свою роботу на інших. Варто проаналізувати причини низької залученості до командних процесів. Розгляньте можливість проведення тренінгів для налагодження контакту з персоналом, підвищення рівня єдності колективу.");
        }else if((averegeTests[0][1]<90)&&(averegeTests[0][1]>=70)){
            setCellValueStringSecondSheet(30,7,"" +
                    "Вам властивий високий рівень комунікації, що допомагає вам легко знаходити спільну мову та активно співправцювати з колегами. Ви повністю розділяєте інтереси команди і часто ставите їх вище власних. Ви добре знаєте посадові обов`язки кожного з підлеглих та слідкуєте за їх чітким дотриманням. Правильно розтавляєте пріоритети, при винекненні будь-якої проблеми сприяєте її швидкому вирішенню. Поглибити свої навики в управлінні та комунікації ви можете звернувшись до відповідної літератури на тему: \"Менеджмент\".");
        }else if (averegeTests[0][1]>=90){
            setCellValueStringSecondSheet(30,7,"" +
                    "Ви однозначно знаєте секрет, як досягти максимальної згуртованості та підняти командний дух в колективі. Завжди влучно формулюєте як загальну, так і особистісну мотивацію, що допомагає залучити до роботи всіх співробітників. Вам завжди вдається правильно розподілити обов'язки та сприяти виконанню всіх необхідних процесів. Ви з азартом беретесь за будь-яке завдання, часто демонструєте надзусилля. Для вас надзвичайно важлива думка та схвалення керівника при виконанні завдань будь-якої складності. Щоразу ставите вищу ціль за попередню та знаходите максимально вірні шляхи для її досягнення.");
        }

        //-------------------3
        if(averegeTests[0][2]<50){
            setCellValueStringSecondSheet(31,7,"" +
                    "Завдання, що ви поручаєте персоналу не завжди можуть відповідати потребам магазину, а також часто незрозумілі персоналу. Не завжди проявляєте прихольність до компанії від чого ваші цілі можуть різнитись з поставленими цілями керівництва. Для вас важко адаптуватись до будь-яких нововведень чи змін, від цього також страждає інформування персоналу з цього приводу. Радимо попрацювати над пильністю та увагою до вирішення відкритих питань. Спробуйте також при прийнятті рішень враховувати та опиратись на увесь діапазон інформації, що ви маєте, так ваші рішення будуть більш раціональними та обґрунтованими.");
        }else if((averegeTests[0][2]<70)&&(averegeTests[0][2]>=50)){
            setCellValueStringSecondSheet(31,7,"" +
                    "Часто будь-які нововедення чи зміни вами важко сприймаються та не співпадаються з вашими поглядами, що може вплинути на якість та своєчасність інформування персоналу з цього приводу. Вам не завжди вдається проконтролювати усі статуси заявок та відкритих питань, аби вирішити їх у необхідний термін. У випадку винекнення стресових ситуацій чи складних питань використовуєте методи, що не завжди співпадають з посадовими інструкціями чи характеристиками, що мають бути притаманні керівнику. Для підвищення ефективності роботи з інформацією радимо застосувати деталізацію завдань: так ви покроково виконуватимете завдання та зменшите відсоток пропущених чи незавершених.");
        }else if((averegeTests[0][2]<90)&&(averegeTests[0][2]>=70)){
            setCellValueStringSecondSheet(31,7,"" +
                    "В роботі з інформацією ви можете зосередитись на головному, бути уважним до брібниць, аргументувати свою точку зору спираючись на розпорядження/правила компанії. Правильно та своєчасно подаєте інформацію для персоналу. Швидко виявляє проблеми та сприяє їх негайному вирішенню. В робочих процесах ви намагаєтесь діяти в рамках посадових інструкцій. Сприймаєте нововведення бренду, але не проявляєте особливої ініціативи для швидкого їх впровадження всередині колективу.  Для розвитку комплексу компетенцій роботи з інформацією радимо спробувати витрачати щоразу менше часу для виконання однотипних задач.");
        }else if (averegeTests[0][2]>=90){
            setCellValueStringSecondSheet(31,7,"" +
                    "Ваші навики роботи з інформацією всестороньо розвинені: за необхідності ви можете варіювати степені опрацювання даних в залежності від встановленого часу.  При постановці задач для персоналу, ви чітко та правильно розставляєте пріоритети для ефективного їх виконання. Яку б задачу ви не ставили перед собою, незважаючи на труднощі та перешкоди, обов`язково доведете її до кінця. Завжди дієте в рамках своїх посадових інструкцій. Прийнятті вами рішення/пропозиції ґрунтуються виключно на фактах.");
        }

        //-------------------4
        if(averegeTests[0][3]<50){
            setCellValueStringSecondSheet(32,7,"" +
                    "У вас спостерігається високий показник незадоволеності в групі: часто ваші погляди розходяться з поглядами керівництва стосовно затрат часу чи виконання завдань, які ви вважаєте неефективними чи непотрібними. Також вам важко пристосуватись до нових умов праці для адаптації нових правил необхідний значний проміжок часу. Вам необхідно звернути увагу, що низька результативность напряму пов`язана з відсутністю конкретних дій з вашої сторони для поліпшення ситуації. Радимо проаналізувати систему формування цілей, побудови плана та інших показників, оперативного реагування на будь-які зміни.");
        }else if((averegeTests[0][3]<70)&&(averegeTests[0][3]>=50)){
            setCellValueStringSecondSheet(32,7,"" +
                    "Ви можете допускати помилки при невірному трактуванні поставлених задач чи правил бренду, тим самим спрямовуючи резерви компанії у невірному напрямку. Можливо вам варто проаналізувати власні сильні та слабкі сторони, щоб визначити більш чіткий шлях у вирішенні задач, не упускаючи дрібниць.");
        }else if((averegeTests[0][3]<90)&&(averegeTests[0][3]>=70)){
            setCellValueStringSecondSheet(32,7,"" +
                    "При виконанні своїх безпосередніх обов`язків ви слідкуєте не лиш за їх якісним виконанням, але і знаходите способи виконати їх швидше, краще та більш якісно. При вірішенні будь-якого завдання ви стаєте безпосереднім його учасником та своїм прикладом демонструєте важливість кожного із елементів завдання. Можливий шлях для удосконалення необхідних компетенцій - ставити складніші задачі, а для їх виконання більш короткий термін.");
        }else if (averegeTests[0][3]>=90){
            setCellValueStringSecondSheet(32,7,"" +
                    "Лояльність та професіоналізм - ваші сильні сторони, як управлінця. Ви добре працюєте з показниками, маєте чудову результативність та ефективність (формуєте переваги завдань в залежності від їх вигоди для компанії). При постановці цілей ви щоразу робите її складнішими, але при цьому приймаєте рішення та розприділяєте пріоритети на основі точних підрахунків. У всіх ваших вчинках виражена пристрасть і відданість до своєї справи, що допомагає націлити персонал до виконання поставлених завдань. Вам властиве не лише вміння правильно подати/пояснити будь-які нововведення, але й уміння їх передбачити.");
        }

        //-----------------final String

        double testsSum = (tests[1][0]+tests[1][1]+tests[1][2]+tests[1][3])/4;
        double testsOwnSum = (testsSelf[1][0]+testsSelf[1][1]+testsSelf[1][2]+testsSelf[1][3])/4;
        double averegeSum = averegeTests[1][0]+averegeTests[1][1]+averegeTests[1][2]+averegeTests[1][3];

        if((testsSum-testsOwnSum)>0.3){
            setCellValueStringSecondSheet(33,4,"\n" +
                    "Вища оцінка по опитувальному бланку ніж оцінка бланку самооцінювання говорить про достатній рівень самокритики. Ви об`єктивно можете оцінити свої знання та знання ваших підлеглих. Навідміну від колективу, ви вважаєте, що недостатньо добре плануєте свою роботу. Радимо розширити свої професійні навики у публічних виступах, це посилить вашу впевненість у собі як у керівнику.");
        }else if((testsOwnSum-testsSum)>0.3){
            setCellValueStringSecondSheet(33,4,"\n" +
                    "Загальний бал по опитувальному бланку нижче за бал по бланку самооцінювання, що говорить про певні розбіжності. У вас добре розвинені навики оперативного і довготривалого планування, але персонал не завжди може поділяти вашу ініціативу, що впливає на швидкість виконання задач. Можливо варто приділити більше уваги чіткості та обгрунтуванню завдань, що ви ставите перед співробітниками. Радимо переглянути /доповнити власний перелік комунікативних методів.");
        }else{
            setCellValueStringSecondSheet(33,4,"\n" +
                    "Бали по вказаним компетенціям однакові згідно результатів опитування персоналу та результатів самооцінювання. Ви у достатньому об`ємі володієте професійними якостями, що необхідні на вашій посаді.  На думку підлеглих, ви вмієте ефективно розподіляти завдання, планувати роботу колективу. У вас добре розвинена професійна та соціальна рефлексія - це допомагає у формуванні та реалізації цілей.  Для вас також важлива внутрішня атмосфера в колективі, що є сильним внутрішнім мотиватором.");
        }

        //----------------------Last Page------------------------------------------------------------------------
        //--------------------------------------------------
        //--------------------------------------------------



        System.out.println(list.length);
        System.out.println(list[0].length);
        System.out.println(name);

        try{
            for(int i=1; i<listSelf.length; i++){
                if(i==1){
                    for(int j=1; j<listSelf[0].length; j++){
                        setCellValueStringFinalSheet(2,j,listSelf[i-1][j-1]);
                    }
                }
                if(listSelf[i][2].equals(name)){
                    for(int j=1; j<listSelf[0].length; j++){
                        setCellValueStringFinalSheet(3,j,listSelf[i][j-1]);
                    }
                }
            }
        }catch (NullPointerException e){
            System.out.println("Error null1");
        }

        int o=1;
        int k=1;
        int countForPageThreePart2 = 0;
        //System.out.println(list[2][15]);
        //System.out.println(list[2][16]);
        //System.out.println(list[2][17]);
        //System.out.println(list[2][0]);
        //try{
            for(o=1; o<list.length; o++){
                if(o==1){
                    countForPageThreePart2++;
                    for(k=1; k<list[0].length; k++){
                        setCellValueStringFinalSheetError(5 + countForPageThreePart2, k, list[o - 1][k - 1]);
                    }
                }
                if(list[o][2].equals(name)){
                    countForPageThreePart2++;
                    for(k=1; k<list[0].length; k++){
                        setCellValueStringFinalSheetError(5+countForPageThreePart2,k,list[o][k-1]);
                    }
                }

            }
        //}catch (NullPointerException e){
            System.out.println("Error null2");
            System.out.println(o);
            System.out.println(k);
        //}
        //setCellValueStringFinalSheet(6,18,list[1][17]);


        fos = new FileOutputStream(exelWorkPath+"\\"+"Бланк обробки даних SS17 "+name+".xlsx");
        finalwb.write(fos);

    }

    public void setCellValueDouble(int row, int cell, double number){
        rowSet = sheetMain1.getRow(row - 1);
        cellSet = rowSet.getCell(cell - 1);
        cellSet.setCellValue(number);
    }

    public void setCellValuePersent(int row, int cell, String number){
        rowSet = sheetMain1.getRow(row - 1);
        cellSet = rowSet.getCell(cell - 1);
        cellSet.setCellValue(number+"%");
    }

    public void setCellValueString(int row, int cell, String str){
        rowSet = sheetMain1.getRow(row - 1);
        cellSet = rowSet.getCell(cell - 1);
        cellSet.setCellValue(str);
    }

    public void setCellValueDoubleSecondSheet(int row, int cell, double number){
        rowSet = sheetMain2.getRow(row - 1);
        cellSet = rowSet.getCell(cell - 1);
        cellSet.setCellValue(number);
    }

    public void setCellValueStringFinalSheet(int row, int cell, String str){
        rowSet = sheetMain3.getRow(row - 1);
        cellSet = rowSet.getCell(cell - 1);
        try{
            cellSet.setCellValue(str);
        }catch (NullPointerException e){

        }
    }

    public void setCellValueStringFinalSheetError(int row, int cell, String str){
        try{
            if(cell-1==0){
                rowSet2 = sheetMain3.createRow(row-1);
            }
            rowSet2 = sheetMain3.getRow(row - 1);
            //cellSet2 = rowSet2.getCell(cell - 1);
            //if(cellSet2==null){
            cellSet2 = rowSet2.createCell(cell-1);
            cellSet2.setCellType(Cell.CELL_TYPE_STRING);
            //}
            //try{
            cellSet2.setCellValue(str);
            //}catch (NullPointerException e){
            //    System.out.println("error here actually...");
            //}
        }catch(NullPointerException e){
            System.out.println(row);
            System.out.println(cell);
        }

    }

    public void setCellValuePersentSecondSheet(int row, int cell, String number){
        rowSet = sheetMain2.getRow(row - 1);
        cellSet = rowSet.getCell(cell - 1);
        cellSet.setCellValue(number+"%");
    }

    public void setCellValueStringSecondSheet(int row, int cell, String str){
        rowSet = sheetMain2.getRow(row - 1);
        cellSet = rowSet.getCell(cell - 1);
        cellSet.setCellValue(str);
    }



    public void setExelSells(double[][] tests, String name) throws IOException{
        wb1.setSheetName(0, name);
        row1 = sheet1.createRow(1);
        row2 = sheet1.createRow(2);
        double sum = 0;
        for(int i=0; i<4; i++){
            cell1 = row1.createCell(i);
            cell2 = row2.createCell(i);
            cell1.setCellValue(tests[0][i]+"%");
            cell2.setCellValue(tests[1][i]);
        }

        fos = new FileOutputStream(exelWorkPath+"\\"+name+".xlsx");
        wb1.write(fos);
        //fos.close();
    }

    public static String getCellType(Cell cell){                //returns the content of a cell in String type
        String result = "";

        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = cell.getDateCellValue().toString();
                } else {
                    result = String.valueOf(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                result = String.valueOf(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                result = String.valueOf(cell.getCellFormula());
                break;
            case Cell.CELL_TYPE_BLANK:
                //System.out.println();
                break;
            default:
        }

        return result;
    }

    public String[] getMainList(){                                                   //Returns list including only names!
        ArrayList<String> onlyNames = new ArrayList<String>();
        boolean flag = false;                                               //flag that will react to same names in the list
        for(int i=1; i<listSelf.length; i++){
            for(int j=0; j<onlyNames.size(); j++){
                if(listSelf[i][2].equals(onlyNames.get(j))){
                    flag=true;                                           //if we already added name to a list, flag is triggered resulting in a skip
                }
            }
            if(flag==false){                                                             //"skip"
                onlyNames.add(listSelf[i][2]);
            }else{
                flag=false;
            }
        }
        String names[]= new String[onlyNames.size()];                       //switch from ArrayList to array because of a return type needed
        for(int j=0; j<onlyNames.size(); j++){
            names[j]=onlyNames.get(j);
        }
        return names;
    }

    public String[][] getStandartList(){
        return list;
    }                       //returns a list of regular answers
    public String[][] getSelfList(){
        return listSelf;
    }                       //returns a list of "self answers"

}


        //i really really need to start commenting this stuff, cant remember what the actual f*ck i was doing in here a year ago...