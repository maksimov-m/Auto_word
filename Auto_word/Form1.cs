using ExcelDataReader;
using System.Data;
using System.Net.Http.Headers;

namespace Auto_word
{
    public partial class Form1 : Form
    {

        string paths_save; 
        string[] paths_bolvanki = new string[0];
        Settings settings;

        public void Init_data(string[]_paths_bolvanki, string _paths_save, string paths_exel)
        {
            paths_save = _paths_save;

            filename = paths_exel;

            Text = paths_exel;

            OpenExelFile(paths_exel);

            foreach (var item in _paths_bolvanki)
            {
                paths_bolvanki = paths_bolvanki.Append(item).ToArray();
            }

        }

        enum Month
        {
            Января,
            Февраля,
            Марта,
            Апреля,
            Мая,
            Июня,
            Июля,
            Августа,
            Сентября,
            Октября,
            Декабря
        }
        public Form1()
        {
            InitializeComponent();
            
            dataGridView2.Columns.Add("Индекс", "Индекс");
            dataGridView2.Columns.Add("ФИО", "ФИО");
            dataGridView2.Columns.Add("Должность", "Должность");
            dataGridView2.Columns.Add("Кафедра", "Кафедра");
            dataGridView2.Columns.AddRange(new DataGridViewColumn[] { new DataGridViewButtonColumn() });
            dataGridView2[4, 0].Value = "Убрать";


        }
        private string filename = string.Empty;

        private DataTableCollection tableCollection = null;

        private void button1_Click(object sender, EventArgs e)
        {

            string date = dateTimePicker1.Value.ToString("dd.MM.yyyy");
            int month = Int32.Parse(date[3..5]);

            date = $"«{date[0..2]}»  {(Month)month}  {date[6..10]}";

            //string[] paths = { "C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые 2\\Болванка 1.docx",
                //"C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые 2\\Болванка 2.docx"};



            var items = new Dictionary<string, List<string>>
            {
                {"<FIO>", new List<string>() },
                {"<POS>", new List<string>() },
                {"<DATE>", new List<string>() },
                {"<DEP>", new List<string>() },
                {"<PROT>", new List<string>() },

            };

            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                items["<FIO>"].Add(dataGridView2.Rows[i].Cells[1].Value.ToString());
                items["<POS>"].Add(dataGridView2.Rows[i].Cells[2].Value.ToString());
                items["<DEP>"].Add(dataGridView2.Rows[i].Cells[3].Value.ToString());
                items["<DATE>"].Add(date);
                items["<PROT>"].Add(textBox2.Text);

            }

            foreach (var item in paths_bolvanki)
            {
                var helper = new WordHelper(item);
                helper.Process(items);
                //helper.threadStart(items);
            }

            string []names_file = new string[0];
            
            foreach (string item in paths_bolvanki)
            {
                names_file = names_file.Append(item[(item.LastIndexOf("\\") + 1)..paths_bolvanki[0].Length]).ToArray();
            }
            string direct = paths_bolvanki[0][0..(paths_bolvanki[0].LastIndexOf("\\") + 1)];

            string[] res_names = new string[0];
            for (int i = 0; i < names_file.Length; i++)
            {
                for (int j = 0; j < dataGridView2.Rows.Count - 1; j++)
                {
                    res_names = res_names.Append($"{direct}{j} {names_file[i]}").ToArray();
                }
                WordHelper.Merge(res_names, $"{direct}Combine {i + 1}", false, $"{direct}template.docx");

                foreach (var item in res_names)
                {
                    try
                    {
                        File.Delete(item);
                    }
                    catch (Exception)
                    {

                        
                    }
                    
                    //Console.WriteLine(item);
                }
                res_names = new string[0];

            }

            

            MessageBox.Show("Работа завершена!");
            //foreach (var item in collection)
            //{

            //}
            //int name = paths_bolvanki[0].LastIndexOf("\\");
            //Console.WriteLine(paths_bolvanki[0][(name + 1) .. paths_bolvanki[0].Length]);
            //string[] outputFileNames = { String.Format(@"C:\Users\Maksim\Desktop\Балванки\Тестовые 2\Combined 1.docx", Guid.NewGuid()),
            //String.Format(@"C:\Users\Maksim\Desktop\Балванки\Тестовые 2\Combined 2.docx", Guid.NewGuid())};
            //string[] names = { "C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые 2\\0 Болванка 1.docx",
            //"C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые 2\\1 Болванка 1.docx"
            //};
            //string[] names2 = {"C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые 2\\0 Болванка 2.docx",
            //"C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые 2\\1 Болванка 2.docx"};

            //WordHelper.Merge(names, outputFileNames[0], false, "C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые 2\\template.docx");
            //WordHelper.Merge(names2, outputFileNames[1], false, "C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые 2\\template.docx");

            //foreach (var item in outputFileNames)
            //{
            //    WordHelper.Merge(names, item, false, "C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые 2\\template.docx");
            //}




        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();

                if (res == DialogResult.OK)
                {
                    filename = openFileDialog1.FileName;

                    Text = filename;

                    OpenExelFile(filename);

                }
                else
                {
                    throw new Exception("Файл не выбран!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenExelFile(string path)
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }

            });


            reader.Close();


            tableCollection = db.Tables;

            foreach (DataTable tabe in tableCollection)
            {
                toolStripComboBox1.Items.Add(tabe.TableName);
            }

            toolStripComboBox1.SelectedIndex = 0;
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable table = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];

            foreach (DataColumn item in table.Columns)
            {
                dataGridView1.Columns.Add(item.ColumnName, item.Caption);
            }


            foreach (DataRow item in table.Rows)
            {
                dataGridView1.Rows.Add(item.ItemArray);
            }
            dataGridView1.Columns.Add("Control", "");

            for (int i = 0; i < dataGridView1.Columns.Count - 1; i++)
            {
                dataGridView1.Rows[i].Cells[dataGridView1.Rows[i].Cells.Count - 1] = new DataGridViewButtonCell() { };
                dataGridView1.Rows[i].Cells[dataGridView1.Rows[i].Cells.Count - 1].Value = "Добавить";
            }

            //for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //{
            //    dataGridView1[dataGridView1.Rows.Count , i].Value = "Добавить";
            //}


        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 4)
                {
                    var a = dataGridView1.Rows[e.RowIndex];
                    a.Cells[a.Cells.Count - 1].Value = "Удалить";
                    dataGridView1.Rows.RemoveAt(e.RowIndex);
                    dataGridView2.Rows.Insert(0, a);
                    dataGridView1.Refresh();
                    dataGridView2.Refresh();
                    //if ((dataGridView2[0, 0].Value as string).Length == 0)
                    //{
                    //    //dataGridView2.Rows[0].Cells[0].Value = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                    //    //dataGridView2.Rows[0].Cells[1].Value = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    //    //dataGridView2.Rows[0].Cells[2].Value = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    //    //dataGridView2.Rows[0].Cells[3].Value = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    //    //dataGridView2.Rows[0].Cells[4].Value = "Убрать";
                    //    //dataGridView2.Rows.Add();
                    //    var a = dataGridView1.Rows[e.RowIndex];
                    //    a.Cells[a.Cells.Count - 1].Value = "Добавить";
                    //    dataGridView1.Rows.RemoveAt(e.RowIndex);
                    //    dataGridView2.Rows.Insert(1, a);
                    //}
                    //else
                    //{
                    //    //dataGridView2.Rows.Add();

                    //    //dataGridView2.Rows[dataGridView2.Rows.Count - 2].Cells[0].Value = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                    //    //dataGridView2.Rows[dataGridView2.Rows.Count - 2].Cells[1].Value = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    //    //dataGridView2.Rows[dataGridView2.Rows.Count - 2].Cells[2].Value = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    //    //dataGridView2.Rows[dataGridView2.Rows.Count - 2].Cells[3].Value = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    //    //dataGridView2.Rows[dataGridView2.Rows.Count - 2].Cells[4].Value = "Убрать";
                    //    var a = dataGridView1.Rows[e.RowIndex];
                    //    a.Cells[a.Cells.Count - 1].Value = "Добавить";
                    //    dataGridView1.Rows.RemoveAt(e.RowIndex);
                    //    dataGridView2.Rows.Insert(1, a);

                    //}
                    //dataGridView1.Rows.RemoveAt(e.RowIndex);
                    //dataGridView1.Refresh();

                }
            }
            catch (Exception)
            {

            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 4)
                {
                    var a = dataGridView2.Rows[e.RowIndex];
                    a.Cells[a.Cells.Count - 1].Value = "Добавить";
                    dataGridView2.Rows.RemoveAt(e.RowIndex);
                    dataGridView1.Rows.Insert(0, a);
                    dataGridView1.Refresh();
                    dataGridView2.Refresh();

                    //if (dataGridView1[0, 0].Value == "")
                    //{
                    //    //dataGridView1.Rows[0].Cells[0].Value = dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString();
                    //    //dataGridView1.Rows[0].Cells[1].Value = dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();
                    //    //dataGridView1.Rows[0].Cells[2].Value = dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();
                    //    //dataGridView1.Rows[0].Cells[3].Value = dataGridView2.Rows[e.RowIndex].Cells[3].Value.ToString();
                    //    //dataGridView1.Rows[0].Cells[4].Value = "Добавить";

                    //}
                    //else
                    //{




                    //    //datagridview1.rows.add(1);

                    //    //dataGridView1.Rows[dataGridView2.Rows.Count - 2].Cells[0].Value = dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString();
                    //    //dataGridView1.Rows[dataGridView2.Rows.Count - 2].Cells[1].Value = dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();
                    //    //dataGridView1.Rows[dataGridView2.Rows.Count - 2].Cells[2].Value = dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();
                    //    //dataGridView1.Rows[dataGridView2.Rows.Count - 2].Cells[3].Value = dataGridView2.Rows[e.RowIndex].Cells[3].Value.ToString();
                    //    //dataGridView1.Rows[dataGridView2.Rows.Count - 2].Cells[4].Value = "Добавить";


                    //}
                    //dataGridView2.Rows.RemoveAt(e.RowIndex);
                    //dataGridView2.Refresh();

                }
            }
            catch (Exception)
            {

            }
        }

        

        

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (settings == null)
            {
               settings = new Settings(this);
            }

            settings.Show();
            
        }

   

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}