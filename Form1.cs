using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows;

using System.Windows.Controls;
using System.Windows.Documents;



namespace EnglishStuding
    
{
        public partial class Form1 : Form
    {
        Random rnd = new Random();
        int Vocab = 250;
        int Size = 10;
        int Quantity = 10;
        const int wordsNumber = 1000;
        //int wordsNumber500 = 500;
        //int wordsNumber1000 = 1000;
        int[] randomIndex=new int[20];
        string[,] Card10 = new string[10,2];
        string[,] Card15 = new string[15,2];
        string[,] Card20 = new string[20,2];
        string[,] Vocabulary = new string[wordsNumber, 2]
 /*10*/           { {"Hello","Привет" }, {"Tree", "Дерево"}, {"Life","Жизнь"}, {"Home", "Дом"}, {"Go", "Идти"}, {"Make","Делать"}, {"Today","Сегодня"}, {"Who","Кто"}, {"You","Ты"}, {"Mother","Мама" },
 /*20*/           {"School","Школа" },{"Friend","Друг"},{"Want","Хотеть"}, {"Pen","Ручка"},{"Water","Вода"}, {"Road","Дорога"}, {"Way","Путь/Способ"}, {"Good","Хорошо"}, {"Bad","Плохо"}, {"Long","Длинный"},
 /*30*/           {"Black","Черный"},{"White","Белый"},{"I","Я"},{"Me","Мне/Я"},{"She","Она"},{"He","Он"},{"They","Они"},{"Know","Знать"},{"Speak","Говорить"},{"Hand","Рука"},
                  {"We","Мы"},{"Think",""},{"Run",""},{"Sky",""},{"Table",""},{"See",""},{"Chair",""},{"Air",""},{"House",""},{"Work",""},
                  {"Green","Зеленый"},{"Grey",""},{"Yellow",""},{"Up",""},{"Down",""},{"Red",""},{"Brown",""},{"Feel",""},{"Be",""},{"Love",""},
                  {"New","Новый"},{"Old",""},{"Young",""},{"Tomorrow",""},{"Day",""},{"Yesterday",""},{"Night",""},{"Evening",""},{"Star",""},{"Wind",""},
                  {"World","Мир"},{"Low",""},{"Down",""},{"Back",""},{"Front",""},{"Before",""},{"Below",""},{"After",""},{"Animal",""},{"Cat",""},
                  {"Dog","Собака"},{"Pet",""},{"Hamster",""},{"Sugar",""},{"Food",""},{"Eat",""},{"It",""},{"Our",""},{"Your",""},{"Where",""},
                  {"Again","Опять"},{"Often",""},{"Hourse",""},{"Rare",""},{"Seldom",""},{"Every",""},{"Come",""},{"Swim",""},{"Stop",""},{"War",""},
/*100*/           {"Short","Короткий"},{"Together",""},{"Keep",""},{"Give",""},{"Power",""},{"Drink",""},{"Cut",""},{"Follow",""},{"Sea",""},{"Lake",""},
                  {"Quick","Быстрый"},{"Fast",""},{"Almost",""},{"Some",""},{"Sometimes",""},{"Time",""},{"Do",""},{"Around",""},{"Their",""},{"Flower",""},
                  {"Sun","Солнце"},{"Doll",""},{"Grass",""},{"Ball",""},{"Father",""},{"Wall",""},{"Brother",""},{"Sister",""},{"Window",""},{"Toy",""},
                  {"Book","Книга"},{"Floor",""},{"Sunny",""},{"Blue",""},{"Read",""},{"Game",""},{"Write",""},{"Bread",""},{"Shop",""},{"Understand",""},
                  {"Light","Свет"},{"Moon",""},{"Sleep",""},{"Cold",""},{"Forest",""},{"Warm",""},{"Hot",""},{"Fox",""},{"Number",""},{"Teacher",""},
                  {"Apple","Яблоко"},{"Bag",""},{"Difficult",""},{"Letter",""},{"Box",""},{"Ugly",""},{"Ship",""},{"Noise",""},{"Look",""},{"Corner",""},
                  {"Many","Много"},{"Few",""},{"But",""},{"All",""},{"Say",""},{"Weather",""},{"If",""},{"Far",""},{"Count",""},{"For",""},
                  {"In","В"},{"At",""},{"About",""},{"On",""},{"Into",""},{"Here",""},{"Above",""},{"Under",""},{"Little",""},{"Greate",""},
                  {"Carrot","Морковь"},{"End",""},{"Desk",""},{"Silver",""},{"Gold",""},{"Country",""},{"Heavens",""},{"Ocean",""},{"Shore",""},{"Or",""},
                  {"When","Когда"},{"Buy",""},{"Coin",""},{"Plant",""},{"River",""},{"Decide",""},{"Line",""},{"Between",""},{"People",""},{"Fly",""},
 /*200*/          {"Boat","Лодка"},{"Car",""},{"Pencil",""},{"Monkey",""},{"Islend",""},{"Stone",""},{"Ground",""},{"Sand",""},{"System",""},{"Paper",""},
                  {"Sheet","Лист"},{"Digit",""},{"Nation",""},{"Mail",""},{"Table",""},{"String",""},{"Line",""},{"City",""},{"Town",""},{"Jump",""},
                  {"Create","Создавать"},{"Brick",""},{"Sit",""},{"Stay",""},{"Put",""},{"Enter",""},{"Exit",""},{"How",""},{"Which",""},{"Than",""},
                  {"That","Этот"},{"Then",""},{"Summer",""},{"Spring",""},{"Winter",""},{"Autom",""},{"Snow",""},{"Rain",""},{"Temperature",""},{"Can",""},
                  {"January","Январь"},{"February",""},{"March",""},{"April",""},{"May",""},{"June",""},{"July","Июль"},{"August",""},{"September",""},{"October",""},
                  {"November","Ноябрь"},{"December",""},{"Year",""},{"Begin",""},{"Fall",""},{"Split","Разбивать"},{"Receive",""},{"Send",""},{"Use",""},{"Habit",""},
                  {"Input","Вводить"},{"Printer",""},{"Button",""},{"Lamp",""},{"Sound",""},{"Circle","Круг"},{"Squire",""},{"Hear",""},{"Triangle",""},{"Fresh",""},
                  {"Money","Деньги"},{"Wish",""},{"Dream","Мечта"},{"Uncle","Дядя"},{"Aunt","Тетя"},{"Grandmother",""},{"Grandfather",""},{"Daughter",""},{"Girl",""},{"Boy",""},
                  {"Woman","Женщина"},{"Female","Женщина"},{"Men","Человек"},{"Male","Мужчина"},{"Person","Человек"},{"Homeless","Бездомный"},{"Army","Армия"},{"Speed","Скорость"},{"Somewhere","Где-то"},{"Horrible","Ужасный"},
                  {"Mark","Оценка/Знак"},{"Mistake","Ошибка"},{"Error","Ошика"},{"Ice","Лёд"},{"Icecream","Мороженое"},{"Sweet","Сладкий"},{"Candy","Конфета"},{"Chokolate","Шоколад"},{"Baby",""},{"Now",""},
 /*300*/          {"Never","Никогда"},{"Always","Всегда"},{"Bird","Птица"},{"Snake","Змея"},{"Lizard","Ящерица"},{"Ant","Муравей"},{"Gippo","Бегемот"},{"Crocodile","Крокодил"},{"Poison","Яд"},{"Mashroom","Гриб"},
                  {"Have","Иметь"},{"Lesson","Урок"},{"Learn","Изучать"},{"Gun","Ружье"},{"Weapon","Оружие"},{"Battle","Битва"},{"Bottle","Бутылка"},{"And","И"},{"With",""},{"From",""},
                  {"Bat","Летучая мышь"},{"Squirrel",""},{"Wolf",""},{"Kitten",""},{"Repeat",""},{"Listen",""},{"This",""},{"Pistol",""},{"Pupil",""},{"Frog",""},
                  {"Fog","Туман"},{"Fork","Вилка"},{"Spoon","Ложка"},{"Like","Нравиться/Как"},{"Cake","Торт"},{"Besquit","Печенье"},{"Log","Бревно"},{"Hat","Шляпа"},{"Leg",""},{"Wrist",""},
                  {"Dark","Темный"},{"Bright","Яркий"},{"Clever","Умный"},{"Cool","Холодный"},{"Smart",""},{"Interesting",""},{"Rabbit",""},{"Chiken",""},{"Cock",""},{"Duck",""},
                  {"Catch","Ловить"},{"Duckling","Утенок"},{"It","Это/Оно"},{"Bite",""},{"Bike",""},{"Base","Основа/База"},{"Beat",""},{"Scream",""},{"Screan",""},{"Display",""},
                  {"Mouse",""},{"Keyboard",""},{"Wire",""},{"Technical",""},{"Large",""},{"Wide",""},{"Effect",""},{"Knoledge",""},{"Tooth",""},{"Teeth",""},
                  {"Eye",""},{"Ear",""},{"Mouth",""},{"Month",""},{"Chest",""},{"Stomach",""},{"Knee",""},{"Thank you",""},{"Excuse",""},{"Remember",""},
                  {"Remind",""},{"Memory",""},{"Brain",""},{"Ability",""},{"Capable","Совместимый"},{"Force","Сила"},{"Equipment","Оборудование"},{"Language",""},{"Foreign",""},{"Common","Общий"},
                  {"Society","Общество"},{"Communication",""},{"Network",""},{"Social",""},{"Hen",""},{"Goat",""},{"Ribbon",""},{"Lemon",""},{"Potato",""},{"Tomato",""},
/*400*/           {"Cocumber",""},{"Garden","Сад"},{"Kindergarden","Детсад"},{"Color",""},{"Taste",""},{"Change",""},{"Test","Пробовать"},{"Try","Пытаться"},{"Attempt","Попытка"},{"Sorry","Сожалеть"},
                  {"Choose",""},{"Violet",""},{"Step",""},{"Mirrow",""},{"Too",""},{"Rose",""},{"Cage",""},{"Cow",""},{"Hundred",""},{"Thousend",""},
                  {"Trousers","Брюки"},{"Umbrella",""},{"Bell",""},{"Child",""},{"Children",""},{"Strong",""},{"Please",""},{"Bring","Приносить"},{"Tiger",""},{"Lion",""},
                  {"Clock",""},{"Hour",""},{"Minute",""},{"Second","Секунда"},{"Brave",""},{"Coward","Трус"},{"Bear",""},{"Weak",""},{"Sharp",""},{"Eagle",""},
                  {"Tail",""},{"Ring",""},{"Coat",""},{"Tame","Ручной"},{"Claw",""},{"Beak",""},{"Bean",""},{"Bee",""},{"Need",""},{"Hunter","Охотник"},
                  {"Pair",""},{"Shoe",""},{"Sofa",""},{"Plate",""},{"Glass",""},{"Future",""},{"Present",""},{"Last",""},{"Past",""},{"Next",""},
                  {"Additional",""},{"Nest",""},{"Part",""},{"Whole",""},{"Hole",""},{"God",""},{"Name",""},{"Surname",""},{"Lastname",""},{"Firstname",""},
                  {"Family",""},{"Mean","Значить"},{"Season","Сезон"},{"Travel","Путешествовать"},{"Job","Работа"},{"Salary","Зарплата"},{"User",""},{"Funny","Забавный"},{"Open",""},{"Close",""},
                  {"Shake",""},{"Lorry",""},{"Driver",""},{"at all","Совсем"},{"Signal","Сигнал"},{"Must","Должен"},{"Ill","Больной"},{"Healthy","Здоровый"},{"Kite","Бумажный змей"},{"Lazy","Ленивый"},
                  {"Fur-Tree","Ель"},{"Get","Получать"},{"Loose","Терять"},{"Find","Находить"},{"Search","Искать"},{"Hide","Прятать"},{"Found out","Выяснять"},{"Add",""},{"Destract",""},{"Delete",""},
 /*500*/          {"Destroy",""},{"Kill",""},{"Burn",""},{"Construct",""},{"Build",""},{"Building",""},{"Late","Поздно"},{"Early","Рано"},{"Nut","Орех"},{"Boil","Варить"},
                  {"Wood",""},{"Wing",""},{"Any","Какой-нибудь"},{"Bread","Хлеб"},{"Much",""},{"Dinner",""},{"Supper",""},{"Breakfest","Завтрак"},{"Kitchen",""},{"Jug","Кувшин"},
                  {"Tea",""},{"Real","Настоящий"},{"Butter",""},{"Class",""},{"Grade",""},{"Door",""},{"Handle",""},{"Entrance",""},{"Berry","Ягода"},{"Leaf",""},
                  {"Jirrafe",""},{"Honey",""},{"Butterfly",""},{"Ladybird",""},{"Else","Еще"},{"Shade",""},{"Shadow",""},{"Fat",""},{"Thin",""},{"Thick",""},
                  {"Maybe",""},{"Poor",""},{"Rich",""},{"Them",""},{"Shoot",""},{"Pull",""},{"Push",""},{"Raise",""},{"Because",""},{"Why","Почему"},
                  {"Distance",""},{"Page",""},{"List",""},{"Report",""},{"Welcome",""},{"Meet",""},{"Meat",""},{"Soon","Скоро"},{"Hair",""},{"Glad","Радоваться"},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
 /*600*/          {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
 /*700*/          {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
/*800*/           {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
/*900*/           {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
                  {"What",""},{"Several",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},{"",""},
 /*1000*/         {"Room",""},{"Bar",""},{"Bath",""},{"Velocity",""},{"King",""},{"Queen",""},{"Pig",""},{"First",""},{"Second",""},{"Third",""}};


        //string[] RussianWords = { "Привет", "Дерево", "Жизнь", "Дом", "Идти", "Делать", "Сегодня", "Кто", "Ты", "Мама",
        //"Школа", "Друг", "Хотеть", "Ручка", "Вода", "Дорога", "Путь", "Хорошо", "Плохо", "Длинный", "Черный", "Белый", "Я", "Меня", 
        //    "Она", "Он", "Они", "Знать", "Говорить", "Рука"};

        public Form1()
        {
            InitializeComponent();
        }
                
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int SizeVocabTemp= comboBox1.SelectedIndex;
            if (SizeVocabTemp == 0)
                Vocab = 250;
            if (SizeVocabTemp == 1)
                Vocab = 500;
            if (SizeVocabTemp == 2)
                Vocab = 1000;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            int SizeIndexTemp = comboBox2.SelectedIndex;
            if (SizeIndexTemp == 0)
                Size = 10;
            if (SizeIndexTemp == 1)
                Size = 15;
            if (SizeIndexTemp == 2)
                Size = 20;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
           int SizeIndexQuantity = comboBox3.SelectedIndex; ;
            if (SizeIndexQuantity == 0)
                Quantity = 10;
            if (SizeIndexQuantity == 1)
                Quantity = 15;
            if (SizeIndexQuantity == 2)
                Quantity = 20;
            if (SizeIndexQuantity == 3)
                Quantity = 25;
            if (SizeIndexQuantity == 4)
                Quantity = 30;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            File.WriteAllText("file.txt", string.Empty);
            for (int counter=0; counter<Quantity; counter++)
            {
                for (int i = 0; i < Size; i++)
                {
                    randomIndex[i] = rnd.Next(0, Vocab);
                }
                if (Size == 10)
                {
                    for (int k = 0; k < Size; k++)
                    {
                        Card10[k, 0] = Vocabulary[randomIndex[k], 0];
                        Card10[k, 1] = Vocabulary[randomIndex[k], 1];
                    }
                    
                    {
                        int k = 0;
                        for (int i = 0; i < Size; i++)
                        {
                            TextWriter tw = new StreamWriter("file.txt", true);
                            tw.Write(Card10[i, 0]);
                            tw.Write("-");
                            tw.WriteLine(Card10[i, 1]);
                            tw.Close();
                            k = i;
                        }
                        TextWriter tw2 = new StreamWriter("file.txt", true);
                        tw2.WriteLine("----------------------------");
                        //tw2.WriteLine(k);
                        tw2.Close();
                    }
                }
                if (Size == 15)
                {
                    for (int k = 0; k < Size; k++)
                    {
                        Card15[k, 0] = Vocabulary[randomIndex[k], 0];
                        Card15[k, 1] = Vocabulary[randomIndex[k], 1];
                    }
                    
                    {
                        for (int i = 0; i < Size; i++)
                        {
                            TextWriter c = new StreamWriter("file.txt", true);
                            c.Write(Card15[i, 0]);
                            c.Write("-");
                            c.WriteLine(Card15[i, 1]);
                            c.Close();
                        }
                        TextWriter c1 = new StreamWriter("file.txt", true);
                        c1.WriteLine("---------------------------");
                        c1.Close();
                    }
                }
                if (Size == 20)
                {
                    for (int k = 0; k < Size; k++)
                    {
                        Card20[k, 0] = Vocabulary[randomIndex[k], 0];
                        Card20[k, 1] = Vocabulary[randomIndex[k], 1];
                    }
                    
                    {
                        for (int i = 0; i < Size; i++)
                        {
                            TextWriter tw = new StreamWriter("file.txt", true);
                            tw.Write(Card20[i, 0]);
                            tw.Write("-");
                            tw.WriteLine(Card20[i, 1]);
                            tw.Close();
                        }
                        TextWriter tw4 = new StreamWriter("file.txt", true);
                        tw4.WriteLine("---------------------------");
                        tw4.Close();
                    }
                }
                
            } 

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //var _filePath = Path.Combine(Directory.GetCurrentDirectory(), "Cards.xls");
            //Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            //if (excel == null)
            //{
            //    System.Windows.MessageBox.Show("Excel is not properly installed!");
            //    return;
            //}
            //Excel.Workbook xlWorkBook;
            //Excel.Worksheet xlWorkSheet;
            //Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(_filePath);
            //Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            //sheet.Close(true, Type.Missing, Type.Missing);
            //excel.Quit();
            //Process.Start(_filePath);



            //Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            //if (xlApp == null)
            //{
            //    System.Windows.MessageBox.Show("Excel is not properly installed!");
            //    return;
            //}
            //Excel.Workbook xlWorkBook;
            //Excel.Worksheet xlWorkSheet;
            //object misValue = System.Reflection.Missing.Value;
            //xlWorkBook = xlApp.Workbooks.Add("Cards.xls");
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //xlWorkSheet.Cells[1, 1] = "ID";
            //xlWorkSheet.Cells[1, 2] = "Name";
            //xlWorkSheet.Cells[2, 1] = "1";
            //xlWorkSheet.Cells[2, 2] = "One";
            //xlWorkSheet.Cells[3, 1] = "2";
            //xlWorkSheet.Cells[3, 2] = "Two";
            //xlWorkBook.SaveAs("Cards.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            //xlWorkBook.Close(true, misValue, misValue);
            //xlApp.Quit();

            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);

            //System.Windows.MessageBox.Show("Excel file created , you can find the file");

        }
    }
}
