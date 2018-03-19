package main;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sun.nio.ch.IOUtil;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
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
        checkForMultyAnswers();
    }

    public void setFinalExel(double[][] tests, double[][] testsSelf, double[][] averegeTests, String name ) throws IOException {
        finalwb = new XSSFWorkbook(new FileInputStream(exelWorkPath+"\\Бланк обробки даних SS18.xlsx"));
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
                    "За результатами оцінки вашого колективу, ви керівник, який досяг влади лише завдяки своїй посаді і керує людьми винятково з цих позицій. Ваша влада поширюється переважно на виробничі відносини і здійснюється за схемою «керівник — підлеглий». На жаль, ваші лідерські якості слабо виражені та мають недостатній вплив для якісного керування колективом. Вам необхідно поспілкуватись з кожним із колег окремо та виявити у чому ваші погляди розходяться. Якщо ви знайдете спільну мову з колегами, то і ваша лідерська позиція стане більш стійкою.");
        }else if((averegeTests[0][0]<75)&&(averegeTests[0][0]>=50)){
            setCellValueStringSecondSheet(29,7,"" +
                    "По своїй натурі ви більш схильні приймати роль виконавця або підлеглого. Зазвичай ви невпевнені у своїх знаннях у тій чи іншій сферах, що значно уповільнює процес прийняття рішень та впливає на їх ефективність. Вам важко знаходити контакт з підлеглими та досягти порозуміння у складних ситуаціях. Спробуйте почати приймати рішення в тих областях, де відмова або невдача не стане критичною для вашої впевненості у собі і у своїх силах. Радимо активно шукати зворотній зв`язок з підлеглими, це допоможе зрозуміти наскільки ваші уявлення про мету та роль в колективі розходиться з уявленням ваших колег.");
        }else if((averegeTests[0][0]<85)&&(averegeTests[0][0]>=75)){
            setCellValueStringSecondSheet(29,7,"" +
                    "Ви дуже вимогливі до себе і ще більш вимогливі до інших, не вибачаєте жодного свого промаху чи помилки, схильні постійно підкреслювати недоліки інших. І хоча це робиться з найкращих спонукань, усе-таки стає причиною конфліктів у силу того, що небагато хто може терпіти систематичне “пиляння”. Радимо вам менше \"напосідати\" на підлеглих та колег, адже ваша думка про правильне/ідеальне може розходитись з думками інших. Спробуйте бути більш відкритим, не завжди ідеальним, але завжди захопленим своєю роботою лідером.");
        }else if((averegeTests[0][0]<90)&&(averegeTests[0][0]>=85)){
            setCellValueStringSecondSheet(29,7,"" +
                    "У керівництві вам не завжди вдається застосовувати неформальні важелі впливу, і для того, щоб підкреслити свою керівну посаду, ви застосовуєте формальні методи впливу. Для досягнення короткочасних завдань - це не так погано. Але для того, щоб бути в першу чергу лідером, радимо вам більше уваги приділити сильним сторонам ваших підлеглих і формувати завдання ґрунтуючись на них. Справжній лідер, який веде за собою, надзвичайно близький та уважний до кожного з його команди. Сила впливу лідерських рис у керівника формує певний стиль керівництва, який може бути більш або менш ефективним у різних управлінських ситуаціях. Закріпити ваші лідерські якості можна за допомогою публічних виступів: підготовлений матеріал тривалістю навіть 5-7 хвилин на ранкових «5-хвилинках» допоможе почувати себе впевненіше як управлінець.");
        }else if((averegeTests[0][0]<95)&&(averegeTests[0][0]>=90)){
            setCellValueStringSecondSheet(29,7,"" +
                    "Ви безумовно є лідером, якого усі сприймають як частину команди, який направляє колектив для досягнення мети. Ви користуєтесь довірою своїх співробітників і чесність в цьому - головна ваша перевага. У стресових ситуаціях ви не втрачаєте самоконтроль, влучно визначаєте методи вирішення будь-якої ситуації. Для удосконалення своїх якостей керівника радимо більшу увагу приділяти зворотному зв`язку від персоналу, щоб передбачити та уникнути можливе виникнення прихованих конфліктів. Також, радимо розвивати навики управлінця, зокрема, шляхом проведення публічних тривалих виступів: зборів/тренінгів.");
        }else if (averegeTests[0][0]>=95){
            setCellValueStringSecondSheet(29,7,"" +
                    "Ви є головним джерелом натхнення і мотивації для усього колективу. Ви безумовно є керівником, якого поважають, а не бояться. Вашому стилю керівництва властива завзятість, що є головним секретом досягнення бажаного результату. Ви завжди дієте організовано та цілеспрямовано в умовах невизначеності. Вам властивий високий рівень самоконтролю та самокритики. Ви безперечно є тим лідером, на якого ваш персонал дивиться в пошуках впевненості та підтримки у складних ситуаціях.");
        }

        //------------------------2
        if(averegeTests[0][1]<50){
            setCellValueStringSecondSheet(30,7,"" +
                    "");
        }else if((averegeTests[0][1]<75)&&(averegeTests[0][1]>=50)){
            setCellValueStringSecondSheet(30,7,"" +
                    "Неформальний лідер. Ви нерідко стаєте ініціатором конфліктних ситуацій і більш сприяєте їх подальшому розвитку, аніж вирішенню. Не вважаєте за потрібне делегувати повноваження. Зазвичай ставите свої інтереси вище за інтереси команди/компанії. Комунікація з колективом знаходиться на низькому рівні та потребує розвитку даної компетенції. Радимо переглянути стиль управління персоналом та звернути увагу на ефективність якісного делегування. Радимо приймати активнішу участь у вирішенні колективних питань та у робочих процесах загалом.");
        }else if((averegeTests[0][1]<85)&&(averegeTests[0][1]>=75)){
            setCellValueStringSecondSheet(30,7,"" +
                    "Зазвичай ви приймаєте участь в командній роботі виключно у присутності безпосереднього керівника. Співробітникам зазвичай важко знайти з вами спільну мову. Часто використовуєте робочий час для вирішення особистих питань. Вам характерна незацікавленість у досягненні цілей, тому часто  ви можете  перекладати свою роботу на інших або відтягувати термін її виконання. Варто проаналізувати причини низького рівня залучення до командних процесів. Розгляньте можливість проведення тренінгів для налагодження контакту з персоналом, підвищення рівня єдності колективу.");
        }else if((averegeTests[0][1]<90)&&(averegeTests[0][1]>=85)){
            setCellValueStringSecondSheet(30,7,"" +
                    "Сила впливу лідерських рис формує певний стиль управління, і деколи стиль який ви обрали виявляється малоефективним у різних ситуаціях. Досягнуті результати оцінюєте за своїми власними мірками, та не завжди враховуєте думку/відношення колег. Аналізуючи наступну задачу радимо більшу увагу приділити залученню колективу, намагайтесь передбачити їх ставлення/реакцію. Пам`ятайте, лідер не створює себе сам, його скеровує на постійний зріст власна команда!");
        }else if((averegeTests[0][1]<95)&&(averegeTests[0][1]>=90)){
            setCellValueStringSecondSheet(30,7,"" +
                    "Вам властивий високий рівень комунікації, що допомагає вам легко знаходити спільну мову та активно співпрацювати з колегами. Ви повністю розділяєте інтереси команди і часто ставите їх вище власних. Ви добре знаєте посадові обов`язки кожного з підлеглих та слідкуєте за їх чітким дотриманням. Правильно розставляєте пріоритети, при виникненні будь-якої проблеми сприяєте її швидкому вирішенню. Поглибити свої навики в управлінні та комунікації ви можете звернувшись до відповідної літератури на тему: \"Менеджмент\", особливу увагу приділивши методикам налагодження контакту, поглибленню довіри/розуміння зі сторони персоналу.");
        }else if (averegeTests[0][1]>=95){
            setCellValueStringSecondSheet(30,7,"" +
                    "Ви однозначно знаєте секрет, як досягти максимальної згуртованості та підняти командний дух в колективі. Завжди влучно формулюєте як загальну, так і особистісну мотивацію, що допомагає залучити до роботи всіх співробітників. Вам завжди вдається правильно розподілити обов'язки та сприяти виконанню всіх необхідних процесів. Ви з азартом беретесь за будь-яке завдання, часто демонструєте надзусилля. Для вас надзвичайно важлива думка та схвалення керівника при виконанні завдань будь-якої складності. Щоразу ставите вищу ціль за попередню та знаходите максимально вірні шляхи для її досягнення.");
        }

        //-------------------3
        if(averegeTests[0][2]<50){
            setCellValueStringSecondSheet(31,7,"" +
                    "");
        }else if((averegeTests[0][2]<75)&&(averegeTests[0][2]>=50)){
            setCellValueStringSecondSheet(31,7,"" +
                    "Завдання, що ви поручаєте персоналу не завжди можуть відповідати потребам магазину, а також часто незрозумілі персоналу. Не завжди проявляєте прихильність до компанії від чого ваші цілі можуть різнитись з поставленими цілями керівництва. Для вас важко адаптуватись до будь-яких нововведень чи змін, від цього також страждає інформування персоналу з цього приводу. Радимо попрацювати над пильністю та увагою до вирішення відкритих питань. Спробуйте також при прийнятті рішень враховувати та опиратись на увесь діапазон інформації, що ви маєте, так ваші рішення будуть більш раціональними та обґрунтованими.");
        }else if((averegeTests[0][2]<85)&&(averegeTests[0][2]>=75)){
            setCellValueStringSecondSheet(31,7,"" +
                    "Часто будь-які нововведення чи зміни вами важко сприймаються та не співпадають з вашими поглядами, що може вплинути на якість та своєчасність інформування персоналу з цього приводу. Вам не завжди вдається проконтролювати усі статуси заявок та відкритих питань, аби вирішити їх у необхідний термін. У випадку виникнення стресових ситуацій чи складних питань використовуєте методи, що не завжди співпадають з посадовими інструкціями чи якостями, що мають бути притаманні керівнику. Для підвищення ефективності роботи з інформацією радимо застосувати деталізацію завдань: так ви покроково виконуватимете завдання та зменшите відсоток пропущених чи незавершених.");
        }else if((averegeTests[0][2]<90)&&(averegeTests[0][2]>=85)){
            setCellValueStringSecondSheet(31,7,"" +
                    "В роботі з інформацією ви можете не вірно трактувати написане і від цього доносити колективу некоректну інформацію. При ознайомленні з  новими розпорядженнями/інструкціями не завжди вивчаєте весь документ, можете переплутати дані, і як висновок, допустити помилок при його виконанні. Радимо більше часу приділяти вивченню нововведень, для якісного ознайомлення з новими методиками - прочитайте надані матеріали декілька разів. При передачі інформації підлеглим, робіть це в усній формі (не зачитуючи з аркушу), але максимально цитуючи написане.");
        }else if((averegeTests[0][2]<95)&&(averegeTests[0][2]>=90)){
            setCellValueStringSecondSheet(31,7,"" +
                    "В роботі з інформацією ви можете зосередитись на головному, бути уважним до дрібниць, аргументувати свою точку зору спираючись на розпорядження/правила компанії. Правильно та своєчасно подаєте інформацію для персоналу. Швидко виявляє проблеми та сприяє їх негайному вирішенню. В робочих процесах ви намагаєтесь діяти в рамках посадових інструкцій. Сприймаєте нововведення бренду, але не проявляєте особливої ініціативи для швидкого їх впровадження всередині колективу.  Для розвитку комплексу компетенцій роботи з інформацією радимо спробувати витрачати щоразу менше часу для виконання однотипних задач.");
        }else if (averegeTests[0][2]>=95){
            setCellValueStringSecondSheet(31,7,"" +
                    "Ваші навики роботи з інформацією всесторонньо розвинені: за необхідності ви можете варіювати степені опрацювання даних в залежності від встановленого часу.  При постановці задач для персоналу, ви чітко та правильно розставляєте пріоритети для ефективного їх виконання. Яку б задачу ви не ставили перед собою, незважаючи на труднощі та перешкоди, обов`язково доведете її до кінця. Завжди дієте в рамках своїх посадових інструкцій. Прийнятті вами рішення/пропозиції ґрунтуються виключно на фактах.");
        }

        //-------------------4
        if(averegeTests[0][3]<50){
            setCellValueStringSecondSheet(32,7,"" +
                    "");
        }else if((averegeTests[0][3]<75)&&(averegeTests[0][3]>=50)){
            setCellValueStringSecondSheet(32,7,"" +
                    "У вас спостерігається високий показник незадоволеності в групі: часто ваші погляди розходяться з поглядами керівництва стосовно затрат часу чи виконання завдань, які ви вважаєте неефективними чи непотрібними. Також вам важко пристосуватись до нових умов праці для адаптації нових правил необхідний значний проміжок часу. Вам необхідно звернути увагу, що низька результативність напряму пов`язана з відсутністю конкретних дій з вашої сторони для поліпшення ситуації. Радимо проаналізувати систему формування цілей, побудови плану та інших показників, оперативного реагування на будь-які зміни.");
        }else if((averegeTests[0][3]<85)&&(averegeTests[0][3]>=75)){
            setCellValueStringSecondSheet(32,7,"" +
                    "Ви можете допускати помилки при невірному трактуванні поставлених задач чи правил бренду, тим самим спрямовуючи резерви компанії у невірному напрямку. Пам`ятайте, для керівника найголовніше, бачити картину в цілому. Радимо вам більше часу приділити аналізу ситуацій в яких ви, або ваша команда допустили помилок, щоб у подальшому зробити висновки базуючись на зібраній інформації та надати переконливі аргументи  при виборі інших методик для вирішення наступної ситуації. Також радимо вам проаналізувати власні сильні та слабкі сторони, щоб визначити більш чіткий шлях у вирішенні задач, не упускаючи дрібниць.");
        }else if((averegeTests[0][3]<90)&&(averegeTests[0][3]>=85)){
            setCellValueStringSecondSheet(32,7,"" +
                    "Виконуючи те чи інше завдання ви щоразу використовуєте перевірені вами методи, у разі внесення будь-яких змін до завдання, вам важко перелаштуватись та звикнути до нововведень. Вам часто не вистачає наполегливості у досягненні мети і в наслідок ви відчуваєте незадоволеність досягнутим. Навіть у випадку позитивно вирішеного питання, ви не відчуваєте повного задоволення від успіху. Через боязнь невдач ви часто обираєте низький рівень ризику і в результаті, можете не досягнути бажаного результату. Вам варто відмовитись від старих методів, спробувавши їх вдосконалити, можливо, це значно полегшить вам роботу за потребуватиме менших затрат часу.");
        }else if((averegeTests[0][3]<95)&&(averegeTests[0][3]>=90)){
            setCellValueStringSecondSheet(32,7,"" +
                    "При виконанні своїх безпосередніх обов`язків ви слідкуєте не лиш за їх якісним виконанням, але і знаходите способи виконати їх швидше, краще та більш якісно. При вирішенні будь-якого завдання ви стаєте безпосереднім його учасником та своїм прикладом демонструєте важливість кожного із елементів завдання. Але методи, що ви застосовуєте не завжди можуть бути найкраще підібрані, у такому випадку ви можете показати результат лиш кількісно, але не якісно. Радимо вам переглянути методики, що ви використовуєте, можливо, необхідно змінити деякі з них для збільшення рівня ефективності. Можливий шлях для удосконалення необхідних компетенцій - ставити складніші задачі, а для їх виконання більш короткий термін.");
        }else if (averegeTests[0][3]>=95){
            setCellValueStringSecondSheet(32,7,"" +
                    "Лояльність та професіоналізм - ваші сильні сторони, як управлінця. Ви добре працюєте з показниками, маєте чудову результативність та ефективність (формуєте переваги завдань в залежності від їх вигоди для компанії). При постановці цілей ви щоразу робите її складнішими, але при цьому приймаєте рішення та розприділяєте пріоритети на основі точних підрахунків. У всіх ваших вчинках виражена пристрасть і відданість до своєї справи, що допомагає націлити персонал до виконання поставлених завдань. Вам властиве не лише вміння правильно подати/пояснити будь-які нововведення, але й уміння їх передбачити.");
        }

        //-----------------final String

        double testsSum = (tests[1][0]+tests[1][1]+tests[1][2]+tests[1][3])/4;
        double testsOwnSum = (testsSelf[1][0]+testsSelf[1][1]+testsSelf[1][2]+testsSelf[1][3])/4;
        double averegeSum = averegeTests[1][0]+averegeTests[1][1]+averegeTests[1][2]+averegeTests[1][3];

        System.out.println("overall point diff: "+(testsSum-testsOwnSum));

        if((testsSum-testsOwnSum)>0.4){
            setCellValueStringSecondSheet(33,4,"\n" +
                    "Оцінка по опитувальному бланку значно вища ніж оцінка бланку самооцінювання говорить про заниження реальних ваших можливостей. Зазвичай це вказує на сильну невпевненість у собі, боязкості та страху перед новими задачами, неможливості реалізувати свої здібності. Вам важко поставити перед собою мету, яку важко досягти, тому ваш вибір - це завдання, що потребують мінімум енергії та ресурсів. Головний ваш страх - зробити помилку, тим паче, що ви займаєте керівну посаду.  Вам варто бути більш впевненішим в собі, навіть якщо ви вирішуєте якесь завдання не вперше, будьте впевненні у вірності свого рішення. І найголовніше - будьте менш критичним до себе яка б ситуація не виникла.");
        }else if(((testsSum-testsOwnSum)<=0.4) && ((testsSum-testsOwnSum)>0.1)){
            setCellValueStringSecondSheet(33,4,"\n" +
                    "Оцінка по опитувальному бланку дещо вища ніж оцінка бланку самооцінювання, що говорить про високий рівень самокритики. Ви об`єктивно можете оцінити свої знання та знання ваших підлеглих. На відміну від колективу, ви вважаєте, що недостатньо добре плануєте свою роботу. Радимо розширити свої професійні навики у публічних виступах, це посилить вашу впевненість у собі, як у керівнику.");
        }else if((testsOwnSum-testsSum)>0.4){
            setCellValueStringSecondSheet(33,4,"\n" +
                    "Загальний бал по опитувальному бланку значно нижче за бал по бланку самооцінювання. Таке співвідношення говорить про неправильне уявлення про себе, ідеалізацію власного образу особистості, своїх можливостей та цінності для навколишніх, свою важливість для вирішення будь-яких питань. В таких випадках людина ігнорує невдачі заради збереження звичної високої оцінки самого себе, своїх вчинків і справ. Відбувається гостре емоційне “відштовхування” усього, що порушує уявлення про себе. Радимо вам не так радикально відноситись до критики, більше прислухатись до зауважень колег, можливо, вони не такі вже і безпідставні та допоможуть вам зробити \"крок уперед\" для налагодження атмосфери в колективі.");
        }else if(((testsOwnSum-testsSum)<=0.4) && ((testsOwnSum-testsSum)>0.1)){
            setCellValueStringSecondSheet(33,4,"\n" +
                    "Загальний бал по опитувальному бланку дещо нижче за бал по бланку самооцінювання, що говорить про певні розбіжності. У вас добре розвинені навики оперативного і довготривалого планування, але персонал не завжди може поділяти вашу ініціативу, що впливає на швидкість виконання задач. Можливо варто приділити більше уваги чіткості та обґрунтуванню завдань, що ви ставите перед співробітниками. Радимо переглянути /доповнити власний перелік комунікативних методів.");
        }else{
            setCellValueStringSecondSheet(33,4,"\n" +
                    "Бали по вказаним компетенціям однакові згідно результатів опитування персоналу та результатів самооцінювання. Ви у достатньому об`ємі володієте професійними якостями, що необхідні на вашій посаді.  На думку підлеглих, ви вмієте ефективно розподіляти завдання, планувати роботу колективу. У вас добре розвинена професійна та соціальна рефлексія - це допомагає у формуванні та реалізації цілей.  Для вас також важлива внутрішня атмосфера в колективі, що є сильним внутрішнім мотиватором.");
        }

        //----------------------Last Page------------------------------------------------------------------------
        //--------------------------------------------------
        //--------------------------------------------------


        CellStyle style = finalwb.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(HSSFCellStyle.FINE_DOTS);
        Font font = finalwb.createFont();
        font.setColor(IndexedColors.YELLOW.getIndex());
        style.setFont(font);

        System.out.println(list.length);
        System.out.println(list[0].length);
        System.out.println(name);

        boolean coloredFlag = false;

        try{
            for(int i=1; i<listSelf.length; i++){
                if(i==1){
                    for(int j=1; j<listSelf[0].length; j++){
                        setCellValueStringFinalSheet(2,j,listSelf[i-1][j-1],coloredFlag,style);
                    }
                }
                if(listSelf[i][2].equals(name)){
                    String[] splitedDate = listSelf[i][0].split(" ");
                    if((Integer.valueOf(splitedDate[2]))>=15){                  //catching people who answered passed the deadline
                        System.out.println(listSelf[i][0]);
                        coloredFlag = true;
                    }
                    for(int j=1; j<listSelf[0].length; j++){
                        setCellValueStringFinalSheet(3,j,listSelf[i][j-1],coloredFlag,style);
                    }
                    coloredFlag = false;
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
                        setCellValueStringFinalSheetError(5 + countForPageThreePart2, k, list[o - 1][k - 1],coloredFlag, style);
                    }
                }
                if(list[o][2].equals(name)){
                    countForPageThreePart2++;
                    //System.out.println(list[o][0]);


                    String[] splitedDate = list[o][0].split(" ");
                    if((Integer.valueOf(splitedDate[2]))>=15){                  //catching people who answered passed the deadline
                        System.out.println(list[o][0]);
                        coloredFlag = true;
                    }

                    for(k=1; k<list[0].length; k++){
                        setCellValueStringFinalSheetError(5+countForPageThreePart2,k,list[o][k-1], coloredFlag, style);
                    }
                    coloredFlag = false;
                }
            }
        //}catch (NullPointerException e){
            System.out.println("Error null2");
            System.out.println(o);
            System.out.println(k);
        //}
        //setCellValueStringFinalSheet(6,18,list[1][17]);


        fos = new FileOutputStream(exelWorkPath+"\\"+"Бланк обробки даних SS18 "+name+".xlsx");
        finalwb.write(fos);

    }

    public void checkForMultyAnswers(){
        ArrayList<String> answerNames = new ArrayList<String>();
        boolean answeredTwiceFlag = false;
        String catchedNameTwice = "";

        try {
            Files.write(Paths.get(exelWorkPath+"\\"+"Answered twice list.txt"), (System.getProperty("line.separator")+job1+"--------------------------------------------------").getBytes(), StandardOpenOption.APPEND);
        } catch (IOException e) {
            e.printStackTrace();
        }

        for(int i=1; i<list.length; i++){
            for(int j=0; j<answerNames.size(); j++){
                if(list[i][3].equals(answerNames.get(j))){
                    answeredTwiceFlag=true;                                           //if we already added name to a list, flag is triggered resulting in a skip and writing down the person who has answered twice
                    catchedNameTwice = answerNames.get(j);
                }
            }

            if(answeredTwiceFlag==false){
                answerNames.add(list[i][3]);
            }else{
                try {
                    Files.write(Paths.get(exelWorkPath+"\\"+"Answered twice list.txt"), (System.getProperty("line.separator")+catchedNameTwice).getBytes(), StandardOpenOption.APPEND);
                }catch (IOException e) {
                    System.out.println("'Answered twice list.txt' file does not exist!");
                }
                answeredTwiceFlag=false;
            }
        }

        answerNames.clear();

        for(int i=1; i<listSelf.length; i++){
            for(int j=0; j<answerNames.size(); j++){
                if(listSelf[i][2].equals(answerNames.get(j))){
                    answeredTwiceFlag=true;                                           //if we already added name to a list, flag is triggered resulting in a skip and writing down the person who has answered twice
                    catchedNameTwice = answerNames.get(j);
                }
            }

            if(answeredTwiceFlag==false){
                answerNames.add(listSelf[i][2]);
            }else{
                try {
                    Files.write(Paths.get(exelWorkPath+"\\"+"Answered twice list.txt"), (System.getProperty("line.separator")+catchedNameTwice+" ----- Multy Self Answer").getBytes(), StandardOpenOption.APPEND);
                }catch (IOException e) {
                    System.out.println("'Answered twice list.txt' file does not exist!");
                }
                answeredTwiceFlag=false;
            }
        }


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

    public void setCellValueStringFinalSheet(int row, int cell, String str, boolean coloredFlag, CellStyle style){
        rowSet = sheetMain3.getRow(row - 1);
        cellSet = rowSet.getCell(cell - 1);
        try{
            cellSet.setCellValue(str);
            if(coloredFlag==true){
                cellSet.setCellStyle(style);
            }
        }catch (NullPointerException e){

        }
    }

    public void setCellValueStringFinalSheetError(int row, int cell, String str, boolean colored, CellStyle style){
        try{
            if(cell-1==0){
                rowSet2 = sheetMain3.createRow(row-1);
            }
            rowSet2 = sheetMain3.getRow(row - 1);

            cellSet2 = rowSet2.createCell(cell-1);
            cellSet2.setCellType(Cell.CELL_TYPE_STRING);

            cellSet2.setCellValue(str);
            if(colored==true){
                cellSet2.setCellStyle(style);
            }
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