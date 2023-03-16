using ExcelDataReader;
using System.Data;

namespace Auto_word
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            dataGridView2.Columns.Add("Индекс", "Индекс");
            dataGridView2.Columns.Add("ФИО", "ФИО");
            dataGridView2.Columns.Add("Должность", "Должность");
            dataGridView2.Columns.Add("Кафедра", "Кафедра");
            dataGridView2.Columns.AddRange(new DataGridViewColumn[] { new DataGridViewButtonColumn() });
            dataGridView2[4, 0].Value = "Delete";

        }
        private string filename = string.Empty;

        private DataTableCollection tableCollection = null;

        private void button1_Click(object sender, EventArgs e)
        {
            string[] paths = { "C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые\\Болванка Бюллетень-нов форма 22-02-2018.docx", 
                "C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые\\болванка ПРОТОКОЛ 22-01-2018.docx",
                "C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые\\болванкаРеш Сов 14-06-2018.doc" };

            

            var items = new Dictionary<string, List<string>>
            {
                {"<FIO>", new List<string>() },
                {"<POS>", new List<string>() },
                {"<DATE>", new List<string>() },
                {"<DEP>", new List<string>() },

            }; 

            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                items["<FIO>"].Add(dataGridView2.Rows[i].Cells[1].Value.ToString());
                items["<POS>"].Add(dataGridView2.Rows[i].Cells[2].Value.ToString());
                items["<DEP>"].Add(dataGridView2.Rows[i].Cells[3].Value.ToString());
                items["<DATE>"].Add(dateTimePicker1.Value.ToString($"dd.MM.yyyy"));
            }

            foreach (var item in paths)
            {
                var helper = new WordHelper(item);
                helper.Process(items);
                //helper.threadStart(items);
            }

            string outputFileName = String.Format(@"C:\Users\Maksim\Desktop\Балванки\Тестовые\Combined.docx", Guid.NewGuid());
            string[] names = { "C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые\\0 Болванка Бюллетень-нов форма 22-02-2018.docx", 
                "C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые\\1 Болванка Бюллетень-нов форма 22-02-2018.docx",
            "C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые\\2 Болванка Бюллетень-нов форма 22-02-2018.docx",
            "C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые\\3 Болванка Бюллетень-нов форма 22-02-2018.docx"};

            WordHelper.Merge(names, outputFileName, false, "C:\\Users\\Maksim\\Desktop\\Балванки\\Тестовые\\template.docx");

            foreach (var item in names)
            {
                File.Delete(item);
            }
            
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();

                if(res == DialogResult.OK)
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

            tableCollection = db.Tables;

            toolStripComboBox1.Items.Clear();

            foreach (DataTable tabe in tableCollection)
            {
                toolStripComboBox1.Items.Add(tabe.TableName);
            }

            toolStripComboBox1.SelectedIndex = 0;
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable table = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];

            dataGridView1.DataSource = table;
            dataGridView1.Columns.AddRange(new DataGridViewColumn[] { new DataGridViewButtonColumn() });
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1[4, i].Value = "Add";
            }
            

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if(e.ColumnIndex == 4)
                {
                    if (dataGridView2[0, 0].Value == "")
                    {
                        dataGridView2.Rows[0].Cells[0].Value = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                        dataGridView2.Rows[0].Cells[1].Value = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                        dataGridView2.Rows[0].Cells[2].Value = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                        dataGridView2.Rows[0].Cells[3].Value = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                        dataGridView2.Rows[0].Cells[4].Value = "Delete";
                        //dataGridView2.Rows.Add();
                    }
                    else
                    {
                        dataGridView2.Rows.Add();

                        dataGridView2.Rows[dataGridView2.Rows.Count - 2].Cells[0].Value = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                        dataGridView2.Rows[dataGridView2.Rows.Count - 2].Cells[1].Value = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                        dataGridView2.Rows[dataGridView2.Rows.Count - 2].Cells[2].Value = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                        dataGridView2.Rows[dataGridView2.Rows.Count - 2].Cells[3].Value = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                        dataGridView2.Rows[dataGridView2.Rows.Count - 2].Cells[4].Value = "Delete";

                        
                    }
                    dataGridView1.Rows.RemoveAt(e.RowIndex);
                    dataGridView1.Refresh();

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
                    if (dataGridView1[0, 0].Value == "")
                    {
                        dataGridView1.Rows[0].Cells[0].Value = dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString();
                        dataGridView1.Rows[0].Cells[1].Value = dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();
                        dataGridView1.Rows[0].Cells[2].Value = dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();
                        dataGridView1.Rows[0].Cells[3].Value = dataGridView2.Rows[e.RowIndex].Cells[3].Value.ToString();
                        dataGridView1.Rows[0].Cells[4].Value = "Add";
                        //dataGridView2.Rows.Add();
                    }
                    else
                    {
                        dataGridView1.Rows.Add(1);

                        dataGridView1.Rows[dataGridView2.Rows.Count - 2].Cells[0].Value = dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString();
                        dataGridView1.Rows[dataGridView2.Rows.Count - 2].Cells[1].Value = dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();
                        dataGridView1.Rows[dataGridView2.Rows.Count - 2].Cells[2].Value = dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();
                        dataGridView1.Rows[dataGridView2.Rows.Count - 2].Cells[3].Value = dataGridView2.Rows[e.RowIndex].Cells[3].Value.ToString();
                        dataGridView1.Rows[dataGridView2.Rows.Count - 2].Cells[4].Value = "Add";


                    }
                    dataGridView2.Rows.RemoveAt(e.RowIndex);
                    dataGridView2.Refresh();

                }
            }
            catch (Exception)
            {

            }
        }
    }
}