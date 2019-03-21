using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Spire.Doc;
using System.IO;

namespace Word
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //Load documents
            Document docTitle = new Document();
            docTitle.LoadFromFile(@"F:\08Github\WORD\Word\SAUSAGE常见问题解答-20181228.docx", FileFormat.Docx);

            string str = docTitle.GetText();
            str = str.Replace("\r", "\r\r");
            string[] sArray = str.Split('\n');
            int i = 0;
            string sHeadline;
            while(true)
            {
                if (IsHeadline(sArray[i].Substring(0, 2)))
                {
                    sHeadline = sArray[i];
                    Directory.CreateDirectory(sHeadline);
                    i++;
                }
                
                while (IsTitle(sArray[i].Substring(0, 2)))
                {
                    MDFile file = new MDFile();

                    file.Title = sArray[i];
                    i++;
                    while(!IsTitle(sArray[i].Substring(0, 2)))
                    {
                        file.Content += sArray[i];
                        i++;
                    }

                    //保存MD文件
                    string sss = file.Title + ".md";
                    sss = sss.Replace(" ", "");
                    sss = sss.Replace("\r", "");
                    sss = sss.Replace("\t", "");
                    sss = sss.Replace("？", "");
                    sss = sss.Replace("，", ",");
                    sss = sss.Replace("、", "");
                    sss = sss.Replace("（", "(");
                    sss = sss.Replace("）", ")");
                    sss = sss.Replace("。", ".");
                    sss = sss.Replace("%", "");
                    sss = sss.Replace(":", "");
                    FileStream fd = new FileStream(sss, FileMode.Create);

                    string str1 = "### " + file.Title;
                    byte[] byteArray1 = System.Text.Encoding.Default.GetBytes(str1);
                    fd.Write(byteArray1, 0, byteArray1.Length);

                    string str2 = "---";
                    byte[] byteArray2 = System.Text.Encoding.Default.GetBytes(str2);
                    fd.Write(byteArray2, 0, byteArray2.Length);

                    if (file.Content != null)
                    {
                        byte[] byteArray3 = System.Text.Encoding.Default.GetBytes(file.Content);
                        fd.Write(byteArray3, 0, byteArray3.Length);
                    }
                    fd.Write(byteArray2, 0, byteArray2.Length);

                    fd.Close();

                }
            }


            //Save and Launch
            docTitle.SaveToFile("Merge.docx", FileFormat.Docx);
            docTitle.SaveToFile("Sample.txt", FileFormat.Txt);
        }

        public bool IsHeadline(string str)
        {
            if (str.Contains("1 ") || str.Contains("2 ") || str.Contains("3 ") || str.Contains("4 ") || str.Contains("5 ") || str.Contains("6 ") || str.Contains("7 ") || str.Contains("8 ") || str.Contains("9 ") || str.Contains("10 "))
                return true;
            else
                return false;
        }
        public bool IsTitle(string str)
        {
            if (str.Contains("1.") || str.Contains("2.") || str.Contains("3.") || str.Contains("4.") || str.Contains("5.") || str.Contains("6.") || str.Contains("7.") || str.Contains("8.") || str.Contains("9.") || str.Contains("10."))
                return true;
            else
                return false;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               
        }

        public class MDFile
        {
            public string Title { get; set; }
            public string Content { get; set; }
        }
        
    }
}
