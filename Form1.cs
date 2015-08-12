using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordDocTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            MessageBox.Show("MEK Vocabulary Test Generator\nDeveloped by Rishi Masand August 2015 for MEK Review");
        }
        
        

        private void button1_Click(object sender, EventArgs e)
        {


            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Word Documents (*.docx)|*.docx";
            ofd.ShowDialog();

            string filePath;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filePath = ofd.FileName;

                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                object miss = System.Reflection.Missing.Value;
                object path = filePath;

                Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
                Document document = application.Documents.Open(path);

                docOutput.Items.Clear();
                listBox1.Items.Clear();

                string totaltext = "";
                for (int i = 0; i < document.Paragraphs.Count; i++)
                {
                    totaltext += " \r\n " + document.Paragraphs[i + 1].Range.Text.ToString();
                    docOutput.Items.Add(document.Paragraphs[i + 1].Range.Text.ToString());


                }

                string theSource = totaltext;
                string[] theStringSeparators = new string[] { "\n" };
                string[] theResult;
                theResult = theSource.Split(theStringSeparators, StringSplitOptions.None);

                string writePath = System.IO.Directory.GetCurrentDirectory();
                Console.Write(writePath);

                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.Append(writePath).Append("\\questionCollection");


                for (int w = 0; w < theResult.Count(); w++)
                {





                    if (theResult[w][1] == 'A' || theResult[w][1] == 'B' || theResult[w][1] == 'C' || theResult[w][1] == 'D' || theResult[w][1] == 'E')
                    {
                        if (theResult[w][1] == 'A')
                        {
                            listBox1.Items.Add(theResult[w - 1]);

                        }
                        if (theResult[w][2] == '.')
                        {
                            listBox1.Items.Add(theResult[w]);

                        }
                        //if (theResult[w][1] == 'E')
                        //{
                        //    listBox1.Items.Add("----");

                        //}
                    }


                }


                ((Microsoft.Office.Interop.Word._Application)application).Quit();
            }


            


        }

        private void docOutput_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

            try
            {
                int m = Int32.Parse(textBox1.Text);
                string writePath = System.IO.Directory.GetCurrentDirectory();
                Console.Write(writePath);

                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.Append(writePath).Append("\\questionCollection.txt");

                System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sb.ToString(), true);
                SaveFile.WriteLine("****");
                SaveFile.WriteLine(m.ToString());
                foreach (var item in listBox1.Items)
                {
                    SaveFile.WriteLine(item);
                }

                SaveFile.Close();

                MessageBox.Show("Test saved!");
            }
            catch (FormatException we)
            {
                //Console.WriteLine(we.Message);
                MessageBox.Show("Test Number must be an Integer");
            }


            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            var document = application.Documents.Add();
            //var paragraph = document.Paragraphs.Add();
            //paragraph.Range.Text = "some text";

            //string filename = "testdoc";
            //application.ActiveDocument.SaveAs(filename, WdSaveFormat.wdFormatDocument);
            //((Microsoft.Office.Interop.Word._Document)document).Close();
            //MessageBox.Show("File saved in Documents");
            //document.Close();

            int fromChapter = Int32.Parse(textBox2.Text);
            int toChapter = Int32.Parse(textBox3.Text);

            int chapDifference = toChapter - fromChapter;

            string writePath = System.IO.Directory.GetCurrentDirectory();
            Console.Write(writePath);

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.Append(writePath).Append("\\questionCollection.txt");

            string[] lines = System.IO.File.ReadAllLines(sb.ToString());

            Random rnd = new Random(DateTime.UtcNow.Millisecond);
            int questionNum = rnd.Next(1, 31);

            Random rndChap = new Random(DateTime.UtcNow.Millisecond);
            int chapNum = rndChap.Next(fromChapter, (toChapter + 1));

            bool correctChapterFound = false;

            int[] usedQuestions = new int[1250];

            bool notUsedBefore = true;

            int numOfQuestions = Int32.Parse(textBox4.Text);

            StringBuilder wordDocBuilder = new StringBuilder();
            wordDocBuilder.Append(textBox5.Text).Append("\nName: ___________________________\n").AppendLine("Class: ___________________________");

            for (int u = 0; u < numOfQuestions; u++)
            {
                chapNum = rndChap.Next(fromChapter, (toChapter + 1));
                for (int v = 0; v < lines.Count(); v++)
                {
                    if (lines[v] == "****")
                    {
                        if (lines[v + 1] == chapNum.ToString())
                        {
                            //MessageBox.Show(lines[v + 1]);
                            
                            questionNum = rnd.Next(1, 30);
                            int numToUse = questionNum * 18;
                            //MessageBox.Show(numToUse.ToString());
                            //MessageBox.Show(lines[v + numToUse + 2]);
                            

                            for (int o = 0; o < usedQuestions.Count(); o++)
                            {
                                if (usedQuestions[o] == v + numToUse + 2)
                                {
                                    notUsedBefore = false;
                                    questionNum = rnd.Next(1, 30);
                                    numToUse = questionNum * 18;
                                    o = 0;
                                }
                                else
                                {
                                    notUsedBefore = true;
                                }
                            }


                            var output = System.Text.RegularExpressions.Regex.Replace(lines[v + numToUse + 2], @"[\d-]", string.Empty);
                            

                                if (notUsedBefore)
                                {
                                    usedQuestions[u] = v + numToUse + 2;
                                    //MessageBox.Show(output);
                                    //MessageBox.Show(lines[v + numToUse + 5]);
                                    //MessageBox.Show(lines[v + numToUse + 8]);
                                    //MessageBox.Show(lines[v + numToUse + 11]);
                                    //MessageBox.Show(lines[v + numToUse + 14]);
                                    //MessageBox.Show(lines[v + numToUse + 17]);

                                    StringBuilder outputBuild = new StringBuilder();
                                    outputBuild.Append((u + 1).ToString()).Append(output);

                                    wordDocBuilder.AppendLine(outputBuild.ToString());
                                    wordDocBuilder.AppendLine(lines[v + numToUse + 5]);
                                    wordDocBuilder.AppendLine(lines[v + numToUse + 8]);
                                    wordDocBuilder.AppendLine(lines[v + numToUse + 11]);
                                    wordDocBuilder.AppendLine(lines[v + numToUse + 14]);
                                    wordDocBuilder.AppendLine(lines[v + numToUse + 17]);

                                }
                                else
                                {
                                    u--;
                                }

                            
                            
                        }
                    }
                }
            }

            document.PageSetup.TextColumns.SetCount(2);

            var paragraph = document.Paragraphs.Add();
            paragraph.Range.Font.Name = "Times New Roman";
            paragraph.Range.Font.Size = 9;
            paragraph.LineSpacing = 1F;
            paragraph.Range.Text = wordDocBuilder.ToString();

            string filename = textBox5.Text;
            
            application.ActiveDocument.SaveAs(filename, WdSaveFormat.wdFormatDocument);
            ((Microsoft.Office.Interop.Word._Document)document).Close();
            MessageBox.Show("File saved in Documents");

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }
    }
}
