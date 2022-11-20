
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel; // подключаем бибилиотеку для эксел
using Word = Microsoft.Office.Interop.Word; // подлючаем для ворд
using System.Xml;

namespace Test1
{
    struct User
    {
        public string name;
        public string last;
        public string gender;
        public int age;
        public string status;
    }

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Files | *.xlsx"; // через опен диалог выбираем файл определенного формата
            if (openFileDialog1.ShowDialog() == DialogResult.OK) // проверка что файл выбран
            {
                var li = ReadExсel(openFileDialog1.FileName); // вызываем функции
                WriteXML(li);
            }
        }

        private List<User> ReadExсel(string filename)
        {
            List<User> li = new List<User>();
            Excel.Application xlApp;

            Excel.Workbook xlWorkBook;

            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            int rw = 0;
            int cl = 1;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            li = new List<User>();
            for (int i = 2; i <= rw; i++)
            {
                User u = new User();

                u.last = (string)(range.Cells[i, 1] as Excel.Range).Value;

                u.name = (string)(range.Cells[i, 2] as Excel.Range).Value;

                u.gender = (string)(range.Cells[i, 3] as Excel.Range).Value;

                u.age = (int)(range.Cells[i, 4] as Excel.Range).Value;

                u.status = (string)(range.Cells[i, 5] as Excel.Range).Value;

                li.Add(u);
            }

            xlWorkBook.Close(true, null, null);

            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);

            Marshal.ReleaseComObject(xlWorkBook);

            Marshal.ReleaseComObject(xlApp);
            return li;
        }
        private void WriteXML(List<User> li)
        {
            XmlTextWriter writer = new XmlTextWriter("test1.xml", System.Text.Encoding.UTF8);

            writer.WriteStartDocument(true);

            writer.Formatting = Formatting.Indented;

            writer.Indentation = 2;

            writer.WriteStartElement("Users");

            foreach (var item in li)

            {
                createnode(item, writer);
            }

            writer.WriteEndElement();

            writer.WriteEndDocument();

            writer.Close();

            MessageBox.Show("XML File created ! ");

        }


        private void createnode(User u, XmlTextWriter writer) // создаем узлы

        {
            writer.WriteStartElement("User");
            writer.WriteStartElement("Last");
            writer.WriteString(u.last.ToString());

            writer.WriteEndElement();
            writer.WriteStartElement("Name");
            writer.WriteString(u.name.ToString());

            writer.WriteEndElement();
            writer.WriteStartElement("Gender");
            writer.WriteString(u.gender.ToString());

            writer.WriteEndElement();
            writer.WriteStartElement("Age");
            writer.WriteString(u.age.ToString());

            writer.WriteEndElement();
            writer.WriteStartElement("Status");
            writer.WriteString(u.status.ToString());
            writer.WriteEndElement();

            writer.WriteEndElement();

        }

        private List<User> ReadXML()
        {
            var li = new List<User>();
            XmlTextReader reader = new XmlTextReader("test1.xml");

            while (reader.ReadToFollowing("User"))
            {
                User user = new User();
                reader.Read();
                reader.ReadStartElement("Last");
                user.last = reader.Value;
                reader.Read();
                reader.ReadEndElement();

                reader.Read();
                reader.ReadStartElement("Name");
                user.name = reader.Value;
                reader.Read();
                reader.ReadEndElement();

                reader.Read();
                reader.ReadStartElement("Gender");
                user.gender = reader.Value;
                reader.Read();
                reader.ReadEndElement();

                reader.Read();
                reader.ReadStartElement("Age");
                user.age = Convert.ToInt32(reader.Value);
                reader.Read();
                reader.ReadEndElement();

                reader.Read();
                reader.ReadStartElement("Status");
                user.status = reader.Value;
                reader.Read();
                reader.ReadEndElement();
                li.Add(user);
            }
            MessageBox.Show("DOCX was creted ! ");
            return li;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Text Files | *.docx";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var li = ReadXML();
                WriteDocx(li, saveFileDialog1.FileName);
            }
        }

        private void WriteDocx(List<User> li, string filename)
        {
            int totalfemale = li.Count(s => s.gender == "ж");
            int totalmale = li.Count(s => s.gender == "м");
            int male3040 = li.Count(s => s.gender == "м" && s.age >= 30 && s.age <= 40);
            int premAcc = li.Count(s => s.status == "премиум");
            int standAcc = li.Count(s => s.status == "стандарт");
            int female30 = li.Count(s => s.gender == "ж" && s.age <= 30);

            Word.Application wApp;

            Word.Range range;

            wApp = new Word.Application();
            wApp.Visible = false;
            var wordDoc = wApp.Documents.Add();
            Word.Paragraph par = wordDoc.Paragraphs.Last;
            par.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            par.Range.Font.Size = 14;
            par.Range.Font.Name = "Times New Roman";
            par.Range.Text = "1. Количество женщина = " + totalfemale + ", мужчин = " + totalmale + "." + Environment.NewLine + "2. Количество мужчин в возрасте 30-40 лет = " + male3040 + "." + Environment.NewLine + "3. Количество стандартных аккунтов = " + standAcc + " и премиум-аккаунтов = " + premAcc + "." + Environment.NewLine + "4. Количество женщин с премиум-аккаунтом в возрасте до 30 лет = " + female30 + ".";
            wordDoc.SaveAs2(filename);
            wordDoc.Close();
            wApp.Quit();
        }
    }
}
