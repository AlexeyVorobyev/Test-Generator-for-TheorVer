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
        public int displaySetting { get; set; }
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

        public static void Generate(int amount,string pathToSave,string JSONpath = @"..\..\Vendor\JSON_Files\Theory.json")
        {
            try
            {
                FileInfo theoryJSON = new FileInfo(Path.GetFullPath(@JSONpath));
                StreamReader theoryJSONReader = theoryJSON.OpenText();
                string JsonResponse = theoryJSONReader.ReadToEnd();

                Root theoryJSONDesirealized = JsonConvert.DeserializeObject<Root>(@JsonResponse);

                List<List<char>> answersMatrix = new List<List<char>>();

                Word.Application word = new Word.Application();
                Word.Document doc = word.Documents.Add();

                Word.Paragraph paragraph;

                Random rnd = new Random();
                for (int i = 0; i < amount; i++)
                {
                    answersMatrix.Add(new List<char>());
                    paragraph = doc.Paragraphs.Add();
                    if (i != 0) paragraph.Range.InsertBreak();
                    paragraph.Range.Text += "Вариант " + (i + 1) + ":\n";

                    for (int j = 0; j < theoryJSONDesirealized.types.Count; j++)
                    {
                        Type type = theoryJSONDesirealized.types[j];

                        int randomFirst = rnd.Next(1, type.tasks.Count + 1);

                        Task task = type.tasks[randomFirst - 1];

                        paragraph = doc.Paragraphs.Add();
                        paragraph.Range.Text = "\n" + (j + 1) + "." + task.text + "\n";

                        int signUnicodeNum = 65;
                        string usedVariants = "";

                        if (task.imagesSource != null) if (task.imagesSource.title != null)
                            {
                                paragraph = doc.Paragraphs.Add();
                                Word.InlineShape inlineShape = paragraph.Range.InlineShapes.AddPicture(Path.GetFullPath(task.imagesSource.title.Replace("\\\\", "\\")));
                            }

                        paragraph = doc.Paragraphs.Add();
                        paragraph.Range.Text = "Варианты ответа:\n";

                        
                        if (task.displaySetting == 1)
                        {
                            if (task.imagesSource != null)
                            {
                                if (!(task.imagesSource.answer != null)) goto mark1;
                                paragraph = doc.Paragraphs.Add();
                                Word.Table tableA = doc.Tables.Add(paragraph.Range, 2, 4);
                                tableA.BottomPadding = 0;
                                int k = 1;
                                while (usedVariants.Length != task.answers.Count)
                                {
                                    int randomSecond = rnd.Next(1, task.answers.Count + 1) - 1;
                                    while (usedVariants.Contains(randomSecond.ToString())) randomSecond = rnd.Next(1, task.answers.Count + 1) - 1;
                                    if (randomSecond == 0) answersMatrix[i].Add((char)signUnicodeNum);
                                    usedVariants += randomSecond.ToString();
                                    if (task.answers[randomSecond] == "source")
                                    {
                                        tableA.Cell(1, k).Range.Text = (char)signUnicodeNum + ":";

                                        string pathImg = "";

                                        foreach (List<object> elem in task.imagesSource.answer) if (Convert.ToInt32(elem[0]) == randomSecond) pathImg = Convert.ToString(elem[1]);


                                        Word.InlineShape inlineShape = tableA.Cell(2, k).Range.InlineShapes.AddPicture(Path.GetFullPath(@pathImg.Replace("\\\\", "\\")));
                                    }
                                    else
                                    {
                                        tableA.Cell(1, k).Range.Text = (char)signUnicodeNum + ": " + task.answers[randomSecond];
                                    }
                                    signUnicodeNum++;
                                    k++;
                                }

                            }
                            else goto mark1;
                            mark1:
                            {
                                paragraph = doc.Paragraphs.Add();
                                Word.Table tableA = doc.Tables.Add(paragraph.Range, 1, 4);
                                tableA.BottomPadding = 0;
                                int k = 1;
                                while (usedVariants.Length != task.answers.Count)
                                {
                                    int randomSecond = rnd.Next(1, task.answers.Count + 1) - 1;
                                    while (usedVariants.Contains(randomSecond.ToString())) randomSecond = rnd.Next(1, task.answers.Count + 1) - 1;
                                    if (randomSecond == 0) answersMatrix[i].Add((char)signUnicodeNum);
                                    usedVariants += randomSecond.ToString();
                                    if (task.answers[randomSecond] == "source")
                                    {
                                        tableA.Cell(1, k).Range.Text = (char)signUnicodeNum + ":";

                                        string pathImg = "";

                                        foreach (List<object> elem in task.imagesSource.answer) if (Convert.ToInt32(elem[0]) == randomSecond) pathImg = Convert.ToString(elem[1]);


                                        Word.InlineShape inlineShape = tableA.Cell(1, k).Range.InlineShapes.AddPicture(Path.GetFullPath(@pathImg.Replace("\\\\", "\\")));
                                    }
                                    else
                                    {
                                        tableA.Cell(1, k).Range.Text = (char)signUnicodeNum + ": " + task.answers[randomSecond];
                                    }
                                    signUnicodeNum++;
                                    k++;
                                }
                            }
                        }
                        else
                        {
                            while (usedVariants.Length != task.answers.Count)
                            {
                                paragraph = doc.Paragraphs.Add();
                                int randomSecond = rnd.Next(1, task.answers.Count + 1) - 1;
                                while (usedVariants.Contains(randomSecond.ToString())) randomSecond = rnd.Next(1, task.answers.Count + 1) - 1;
                                if (randomSecond == 0) answersMatrix[i].Add((char)signUnicodeNum);
                                usedVariants += randomSecond.ToString();
                                if (task.answers[randomSecond] == "source")
                                {
                                    paragraph.Range.Text = (char)signUnicodeNum + ": \n";
                                    paragraph = doc.Paragraphs.Add();

                                    string pathImg = "";

                                    foreach (List<object> elem in task.imagesSource.answer) if (Convert.ToInt32(elem[0]) == randomSecond) pathImg = Convert.ToString(elem[1]);


                                    Word.InlineShape inlineShape = paragraph.Range.InlineShapes.AddPicture(Path.GetFullPath(@pathImg.Replace("\\\\", "\\")));
                                }
                                else
                                {
                                    paragraph.Range.Text = (char)signUnicodeNum + ": " + task.answers[randomSecond] + "\n";
                                }
                                signUnicodeNum++;
                            }
                        }
                    }
                }
                doc.SaveAs2(@pathToSave + @"\test.docx");
                doc.Close();

                Word.Document doc2 = word.Documents.Add();

                paragraph = doc2.Paragraphs.Add();
                paragraph.Range.Text = "Столбец - номер задания, Строчка - вариант\n";

                paragraph = doc2.Paragraphs.Add();
                Word.Table table = doc2.Tables.Add(paragraph.Range, answersMatrix.Count + 1, answersMatrix[0].Count + 1);
                table.Borders.Enable = 1;

                for (int i = 0; i < answersMatrix.Count; i++) table.Cell(i + 2, 1).Range.Text = Convert.ToString(i + 1);
                for (int i = 0; i < answersMatrix[0].Count; i++) table.Cell(1, i + 2).Range.Text = Convert.ToString(i + 1);

                for (int i = 0; i < answersMatrix.Count; i++)
                {
                    for (int j = 0; j < answersMatrix[0].Count; j++)
                    {
                        table.Cell(i + 2, j + 2).Range.Text = Convert.ToString(answersMatrix[i][j]);
                    }
                }

                doc2.SaveAs2(@pathToSave + @"\answers.docx");
                doc2.Close();
                MessageBox.Show("Генерация выполнена успешно!", "Mission completed", MessageBoxButtons.OK);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Возникла проблема при генерации теста\n Ошибка: " + ex.Message, "Ошибка", MessageBoxButtons.OK);
            }
        }
    }
}
