using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using Newtonsoft.Json;
using System.Numerics;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace TeorVer
{
    public class ImagesSource
    {
        public List<List<object>> answer { get; set; }
        public string title { get; set; }
    }

    public class Root
    {
        public List<Type> types { get; set; }
    }

    public class Task
    {
        public string text { get; set; }
        public List<string> answers { get; set; }
        public ImagesSource imagesSource { get; set; }
    }

    public class Type
    {
        public List<Task> tasks { get; set; }
    }


    public static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }


        public static void Generate(int amount,string pathToSave)
        {
            FileInfo theoryJSON = new FileInfo(@"..\..\Vendor\JSON_Files\Theory.json");
            StreamReader theoryJSONReader = theoryJSON.OpenText();
            string JsonResponse = theoryJSONReader.ReadToEnd();

            Root theoryJSONDesirealized = JsonConvert.DeserializeObject<Root>(@JsonResponse);

            List<List<List<char>>> answersMatrix = new List<List<List<char>>>(); 

            Word.Application word = new Word.Application();
            Word.Document doc = word.Documents.Add();
            Word.Paragraph paragraph;

            Random rnd = new Random();
            for (int i = 0; i < amount; i++)
            {
                answersMatrix.Add(new List<List<char>>());
                paragraph = doc.Paragraphs.Add();
                if (i != 0) paragraph.Range.InsertBreak();
                paragraph.Range.Text += "Вариант " + (i + 1) + ":\n";
               
                Console.WriteLine("Вариант " + (i + 1) + ":");

                for (int j = 0; j < theoryJSONDesirealized.types.Count;j++)
                {
                    answersMatrix[i].Add(new List<char>());
                    Type type = theoryJSONDesirealized.types[j];

                    
                    Console.Write((j+1) + ". ");

                    int randomFirst = rnd.Next(1, type.tasks.Count + 1);

                    Task task = type.tasks[randomFirst - 1];

                    paragraph = doc.Paragraphs.Add();
                    paragraph.Range.Text = "\n" + (j + 1) + "." + task.text + "\nВарианты ответа:\n";

                    //Console.WriteLine(task.text);
                    //Console.Write("Варианты ответа: \n");

                    int signUnicodeNum = 65;
                    string usedVariants = "";

                    while (usedVariants.Length != task.answers.Count)
                    {
                        paragraph = doc.Paragraphs.Add();
                        int randomSecond = rnd.Next(1, task.answers.Count + 1) - 1;
                        while (usedVariants.Contains(randomSecond.ToString())) randomSecond = rnd.Next(1, task.answers.Count + 1) - 1;
                        if (randomSecond == 0) answersMatrix[i][j].Add((char)signUnicodeNum);
                        usedVariants += randomSecond.ToString();
                        if (task.answers[randomSecond] == "source")
                        {
                            //paragraph.Range.InsertFile(Path.GetFullPath(@"..\..\Vendor\Images\17.1.PNG"));
                            paragraph.Range.Text = (char)signUnicodeNum + ": \n";
                            paragraph = doc.Paragraphs.Add();

                            string pathImg = "";

                            foreach (List<object> elem in task.imagesSource.answer) if (Convert.ToInt32(elem[0]) == randomSecond) pathImg = Convert.ToString(elem[1]);
                           

                            Word.InlineShape inlineShape = paragraph.Range.InlineShapes.AddPicture(Path.GetFullPath(@pathImg.Replace("\\\\","\\")));
                            //inlineShape.Width = 100;
                            //inlineShape.Height = 10;
                        } else
                        {
                            paragraph.Range.Text = (char)signUnicodeNum + ": " + task.answers[randomSecond] + "\n";
                            Console.Write((char)signUnicodeNum + ": " + task.answers[randomSecond] + "  ");
                        }
                        signUnicodeNum++;
                    }

                    Console.Write("\n");
                }
            }
            //doc.SaveAs2(@pathToSave + @"\test.docx");
            doc.SaveAs2(@"C:\Users\miste\OneDrive\Рабочий стол\WordTest\test2.docx");
            doc.Close();
        }
    }
}
