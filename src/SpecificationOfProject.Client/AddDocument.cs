﻿using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DataBaseLibrary;

namespace SpecificationOfProject.Client
{
    public partial class AddDocument : Form
    {
        public AddDocument()
        {
            InitializeComponent();
            // Пока пути пусты, кнопку выключим изначально
            button1.Enabled = false;
        }

        public string projectDocDirectoryPath;
        public string selectedProject;    
        string[] splittedFileName;

        // Кнопка выбора файла для добавления
        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.Multiselect = false;
            //openFileDialog1.ShowDialog();
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            textBox1.Text = openFileDialog1.FileName;
            splittedFileName = openFileDialog1.FileName.Split('\\');
            textBox2.Text = splittedFileName.Last();
            if (comboBox1.Text != "")
            {
                button1.Enabled = true;
            }
        }

        // Кнопка отменить
        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            this.Owner.Show();
            this.textBox1.Clear();
            this.textBox2.Clear();
            this.button1.Enabled = false;
            this.comboBox1.SelectedIndex = -1;
        }

        // Кнопка добавить
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (DataBaseContext dataBaseConnection = new DataBaseContext())
                {
                    // Пробую получить хоть какие данные для проверки доступности БД (костыль)
                    var testConnect = dataBaseConnection.Projs.LongCount().ToString();

                    // Записываю данные для записи
                    var documentForProject = new DocumentForProject();
                    documentForProject.ProjectID = Convert.ToInt32(selectedProject);
                    documentForProject.DocumentName = splittedFileName.Last();
                    documentForProject.DocumentType = comboBox1.Text;

                    // Копирую файл в директорию документов проекта
                    var filePath = @"" + openFileDialog1.FileName;
                    var fileInfo = new FileInfo(filePath);
                    var newDocPath = projectDocDirectoryPath + "\\" + splittedFileName.Last();
                    if (fileInfo.Exists == true)
                    {
                        fileInfo.CopyTo(newDocPath, true);
                    }
                    else
                    {
                        var fileIsNotExist = new Exception("Ошибка копирования, проблемы с файлом.");
                        throw fileIsNotExist;
                    }
                    documentForProject.DocumentPath = newDocPath;

                    // Вношу изменения в базу данных
                    dataBaseConnection.DocumentForProjects.Add(documentForProject);
                    dataBaseConnection.SaveChanges();
                }
                // Закрываю форму и очищаю поля
                button2_Click(null, null);
                // Указал, что эта форма подчиняется главной
                var mainForm = this.Owner as MainForm;
                // Вызвал функцию в главной форме на обновление грида
                mainForm.FillDocumentsForProjectGrid();
            }
            catch
            {
                MessageBox.Show("Ошибка при добавлении документа к проекту\nПроверьте БД (ошибка в функции добавления)");
            }          
        }

        // Проверяю индекс комбобокса и данные в текстбоксах для активации кнопки
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (textBox1.Text != string.Empty && textBox2.Text != string.Empty && comboBox1.Text != string.Empty)
            {
                button1.Enabled = true;
            }
        }
    }
}
