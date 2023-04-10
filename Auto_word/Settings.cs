using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Auto_word
{
    public partial class Settings : Form
    {
        string[] paths_bolvanki = new string[0];
        Form1 form; 

        public Settings(Form1 form)
        {
            InitializeComponent();
            this.form = form;   
        }


        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();
            //Console.WriteLine(folderBrowserDialog.SelectedPath);
            textBox1.Text = folderBrowserDialog.SelectedPath;
        }

        private void Settings_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            //openFileDialog.ShowDialog(); 
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "DOCX(*.docx)|*.docx|DOC(*.doc)|*.doc";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                comboBox1.Items.Clear();
                paths_bolvanki = openFileDialog.FileNames;
                foreach (var item in paths_bolvanki)
                {
                    comboBox1.Items.Add(item);

                    //Console.WriteLine(item);

                }

                comboBox1.SelectedItem = paths_bolvanki[0];
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();

                if (res == DialogResult.OK)
                {
                    textBox2.Text = openFileDialog1.FileName;

                    //Text = filename;

                    //OpenExelFile(filename);

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

        private void button4_Click(object sender, EventArgs e)
        {
            if(textBox1.Text != "" && textBox2.Text != "" && comboBox1.Items.Count != 0)
            {
                form.Init_data(paths_bolvanki, textBox1.Text, textBox2.Text);

                this.Close();
            }
            else
            {
                MessageBox.Show("Заполните все данные!");
            }
        }
    }
}
