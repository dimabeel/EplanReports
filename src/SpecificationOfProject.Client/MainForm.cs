using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DataBaseLibrary;
using Component = DataBaseLibrary.Component;

namespace SpecificationOfProject.Client
{

    public partial class MainForm : Form
    {
        // Обращаемся к INI файлу
        ConfigIniFile iniFile = new ConfigIniFile("SpecificationOfProject.ClientConfig.ini");
        // Путь хранения документов
        public string documentsPath;
        // Выбранный проект
        public string selectedProject;
        // Форма добавления файла
        AddDocument addDocument = new AddDocument();

        public MainForm()
        {
            InitializeComponent();
            
            // Определили владельца формы
            addDocument.Owner = this;
            
            //Проверяем ini файл клиента
            chekClientIniFile();
            
            // Заполняем дерево при запуске
            FillTreeview();
        }

        // Проверяем ini файл клиента с настройками
        private void chekClientIniFile()
        {
            try
            {
                if (iniFile.KeyExists("Path", "ProjectDocuments"))
                {
                    var iniProperty = iniFile.ReadINI("ProjectDocuments", "Path");
                    documentsPath = @"" + Convert.ToString(iniProperty);
                }
                else
                {
                    MessageBox.Show("Проверьте настройки конфигурационного файла клиента, некорректные настройки");
                    Application.Exit();
                }
            }
            catch
            {
                MessageBox.Show("Проверьте конфигурационноый ini файл клиента");
            }   
        }

        // Выход из клиента
        private void button4_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        // Обновить дерево
        private void button3_Click(object sender, EventArgs e)
        {
            FillTreeview();
            label3.Text = "не выбран";
        }

        // Заполнить дерево
        public void FillTreeview()
        {
            // Отключаем кнопки т.к дерево обновляется и проект не выбран
            button1.Visible = false;
            button2.Visible = false;
            button5.Visible = false;
            button7.Visible = false;
            button6.Visible = false;
            // Очищаем грид перед заполнением
            dataGridView1.Columns.Clear();
            dataGridView1.Refresh();
            // Очищаем перед заполнением дерево
            treeView1.Nodes.Clear();
            treeView1.Refresh();
            try
            {
                using (DBContext DBCon = new DBContext())
                {
                    // Получаем проекты
                    Proj[] projs = DBCon.Projs.ToArray();

                    // Перебираем проекты
                    foreach (Proj proj in projs)
                    {
                        // Получаю изделия по проекту
                        PArticle[] pArticles = DBCon.PArticles.Where(
                            o => o.ProjectID == proj.ProjID).ToArray();
                        // Создаю treeNode с  нулевым уровнем = proj.Name
                        TreeNode projNode = new TreeNode(proj.Name);
                        projNode.Name = proj.ProjID.ToString();
                        projNode.Text = proj.Name;
                        foreach (PArticle pArticle in pArticles)
                        {
                            // Ищу описание по изделию
                            string descr = DBCon.LocationDescriptions.Where(
                                o => o.LocationDescriptionID == pArticle.LocationDesriptionID).
                                Select(o1 => o1.Description).FirstOrDefault();
                            string locationID = DBCon.LocationDescriptions.Where(
                                o => o.LocationDescriptionID == pArticle.LocationDesriptionID).
                                Select(o1 => o1.LocationDescriptionID).FirstOrDefault().ToString();
                            // Заполняю первый уровень
                            projNode.Nodes.Add(locationID, descr);
                            // Заполняю второй уровень           
                            List<Component> components = DBCon.Components.Where(
                                o => o.PArticleID == pArticle.PArticleID).ToList();
                            TreeNode articleNode = new TreeNode();
                            foreach (Component component in components)
                            {
                                projNode.LastNode.Nodes.Add(component.PartNumber, component.PartNumber);
                            }
                        }
                        treeView1.Nodes.Add(projNode);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Ошибка подключения к БД (функция FillTreeView)");
            }     
        }

        // Событие после нажатия на элемент дерева
        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            // Видимость кнопок после выбора какого-либо элемента в дереве
            if (treeView1.SelectedNode != null)
            {
                button1.Visible = false;
                button2.Visible = false;
                button5.Visible = true;
                button6.Visible = false;
                button7.Visible = false;
                label3.Text = "не выбран";
            }

            const int minComponentNameLength = 3;
            const int minComponentDescriptionLength = 5;
            FunctionManager FM = new FunctionManager();
            // Если выбрано значение в дереве и оно первого уровня вложенности
            if (treeView1.SelectedNode != null && treeView1.SelectedNode.Level == 1)
            {
                // Записываю, какой проект выбран сейчас (ProjID)
                selectedProject = treeView1.SelectedNode.Parent.Name;
                label3.Text = treeView1.SelectedNode.Parent.Text;
                string selectedItemParentName = treeView1.SelectedNode.Parent.Text;
                int selectedItemID = Convert.ToInt32(treeView1.SelectedNode.Name);
                try
                {
                    // Если к Бд подключилс, то и настраивай колонки
                    using (DBContext DBCon = new DBContext())
                    {
                        // Настройка колонок первого уровня
                        setUpFirstLevelGridColumns();
                        // Нашел проект к которому принадлежит узел в дереве
                        Proj proj = DBCon.Projs.Where(
                            o => o.Name == selectedItemParentName).FirstOrDefault();
                        // Нашел код описания изделия
                        int selectedItemLocationID = DBCon.LocationDescriptions.Where(
                            o => o.LocationDescriptionID == selectedItemID).Select(
                            o1 => o1.LocationDescriptionID).FirstOrDefault();
                        // Нашел изделие, которое выбрано в дереве по коду
                        PArticle pArticle = DBCon.PArticles.Where(
                            o => o.ProjectID == proj.ProjID).Where(
                            o1 => o1.LocationDesriptionID == selectedItemLocationID).FirstOrDefault();
                        // Нашел все компоненты по изделию
                        List<Component> components = DBCon.Components.Where(
                            o => o.PArticleID == pArticle.PArticleID).ToList();
                        // Перебираю компоненты, заполняю грид
                        foreach (Component component in components)
                        {
                            string componentType = DBCon.ComponentCatalogs.Where(
                                o => o.PartNumber == component.PartNumber).Select(
                                o1 => o1.TypeNumber).FirstOrDefault();
                            string componentManufacturer = DBCon.ComponentCatalogs.Where(
                                o => o.PartNumber == component.PartNumber).Select(
                                o1 => o1.ManufacturerFullName).FirstOrDefault();
                            string componentDescr1 = DBCon.ComponentCatalogs.Where(
                                o => o.PartNumber == component.PartNumber).Select(
                                o1 => o1.Description1).FirstOrDefault();
                            string componentDescr2 = DBCon.ComponentCatalogs.Where(
                                o => o.PartNumber == component.PartNumber).Select(
                                o1 => o1.Description2).FirstOrDefault();
                            string componentInfo = componentManufacturer + " " + componentType;
                            string componentDescription = componentDescr1 + " (" + componentDescr2 + ")";
                            if (componentInfo.Length < minComponentNameLength) componentInfo = component.PartNumber;
                            if (componentDescription.Length < minComponentDescriptionLength) componentDescription = "Отсутствует";
                            dataGridView1.Rows.Add(componentInfo, componentDescription, component.Count);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("Ошибка в БД (функция treeView1_AfterSelect level 1)");
                }
                dataGridView1.ClearSelection();
            }

            // Если выбрано значение в дереве и оно нулевого уровня вложенности
            if (treeView1.SelectedNode != null && treeView1.SelectedNode.Level == 0)
            {
                // Записываю, какой проект выбран сейчас (ProjID)
                selectedProject = treeView1.SelectedNode.Name;
                label3.Text = treeView1.SelectedNode.Text;
                string selectedItem = treeView1.SelectedNode.Text;
                try
                {
                    // Если к Бд подключилс, то и настраивай колонки
                    using (DBContext DBCon = new DBContext())
                    {
                        // Настраиваю колонки нулевого уровня
                        setUpZeroLevelGridColumns();
                        // Нашел проект к которому принадлежит узел в дереве
                        Proj proj = DBCon.Projs.Where(
                            o => o.Name == selectedItem).FirstOrDefault();
                        // Нашел все изделия в проекте
                        List<PArticle> pArticles = DBCon.PArticles.Where(
                            o => o.ProjectID == proj.ProjID).ToList();
                        // Заполняю грид
                        foreach (PArticle pArticle in pArticles)
                        {
                            string descr = DBCon.LocationDescriptions.Where(
                                o => o.LocationDescriptionID == pArticle.LocationDesriptionID).Select(
                                o1 => o1.Description).FirstOrDefault();
                            int count = 0;
                            List<Component> components = DBCon.Components.Where(
                                o => o.PArticleID == pArticle.PArticleID).ToList();
                            foreach (Component component in components)
                            {
                                count += component.Count;
                            }
                            dataGridView1.Rows.Add(descr, count);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("Ошибка в БД (функция treeView1_AfterSelect level 0)");
                }
                dataGridView1.ClearSelection();
            }

            // Если выбрано значение в дереве и оно второго уровня вложенности
            if (treeView1.SelectedNode != null && treeView1.SelectedNode.Level == 2)
            {
                // Записываю, какой проект выбран сейчас (ProjID)
                selectedProject = treeView1.SelectedNode.Parent.Parent.Name;
                label3.Text = treeView1.SelectedNode.Parent.Parent.Text;
                // Ищу выбранный компонент и по нему заполняю грид
                string componentName = treeView1.SelectedNode.Name;
                List<ComponentPropertiesInfo> componentPropertiesInfos = FM.GetPropertiesInfos(componentName);
                // Настраиваю колонки второго уровня              
                setUpSecondLevelGridColumns();
                foreach (ComponentPropertiesInfo componentPropertiesInfo in componentPropertiesInfos)
                {
                    if ((componentPropertiesInfo.Value != "") &&
                        (componentPropertiesInfo.Value != "0") &&
                        (componentPropertiesInfo.Value != null))
                    {
                        dataGridView1.Rows.Add(componentPropertiesInfo.Property, componentPropertiesInfo.Value);
                    }
                }
                dataGridView1.ClearSelection();
            }
        }

        // Документация по проекту
        private void button5_Click(object sender, EventArgs e) 
        {
            // Настраиваю колонки, включаю кнопки
            setUpProjectDocumentsGridColumns();

            // Проверяю, есть ли каталог проекта в хранилище.
            string projectPath = documentsPath + selectedProject.ToString();
            if (Directory.Exists(projectPath))
            {
                FillDocumentsForProjectGrid();
            }
            else
            {
                // Если каталога нету, создаю его
                Directory.CreateDirectory(projectPath);
            }

            button1.Visible = true;
            button2.Visible = true;
            button6.Visible = true;
            button7.Visible = true;
        }

        // Сформировать спецификацию
        private void button2_Click(object sender, EventArgs e)
        {
            FunctionManager FM = new FunctionManager();
            string savePath = documentsPath + selectedProject;
            string projectName = label3.Text;
            FM.CreateProjectSpecification(selectedProject, savePath, projectName);
            try
            {
                // Сохраняю в БД
                using (DBContext DBCon = new DBContext())
                {
                    // Проверяю, есть ли такой документ по проекту
                    int projectID = Convert.ToInt32(selectedProject);
                    string docName = DBCon.DocumentForProjects.Where(o => o.ProjectID == projectID).Where(
                        o1 => o1.DocumentName == "Спецификация.xlsx").Select(
                        o2 => o2.DocumentName).FirstOrDefault();
                    // Если документа нет, сохраняем, если есть - он перезапишется сам
                    if (docName == null)
                    {
                        // Записываю данные для записи
                        DocumentForProject documentForProject = new DocumentForProject();
                        documentForProject.ProjectID = Convert.ToInt32(selectedProject);
                        documentForProject.DocumentName = "Спецификация.xlsx";
                        documentForProject.DocumentType = "Спецификация";
                        documentForProject.DocumentPath = savePath + "\\" + documentForProject.DocumentName;
                        // Вношу изменения в базу данных
                        DBCon.DocumentForProjects.Add(documentForProject);
                        DBCon.SaveChanges();
                    }
                }
                // Обновляю грид
                FillDocumentsForProjectGrid();
            }
            catch
            {
                MessageBox.Show("Ошибка в сохранении документа <Спецификация> в БД, проверьте ее\nДокумент создан, но не внесен в БД");
            }
        }

        // Сформировать заявку на склад
        private void button1_Click(object sender, EventArgs e)
        {
            FunctionManager FM = new FunctionManager();
            string savePath = documentsPath + selectedProject;
            FM.CreateProjectWarehouseRequest(selectedProject, savePath);
            try
            {
                // Сохраняю в БД
                using (DBContext DBCon = new DBContext())
                {
                    // Проверяю, есть ли такой документ по проекту
                    int projectID = Convert.ToInt32(selectedProject);
                    string docName = DBCon.DocumentForProjects.Where(o => o.ProjectID == projectID).Where(
                        o1 => o1.DocumentName == "Заявка на склад.xlsx").Select(
                        o2 => o2.DocumentName).FirstOrDefault();
                    // Если документа нет, сохраняем, если есть - он перезапишется сам
                    if (docName == null)
                    {
                        // Записываю данные для записи
                        DocumentForProject documentForProject = new DocumentForProject();
                        documentForProject.ProjectID = Convert.ToInt32(selectedProject);
                        documentForProject.DocumentName = "Заявка на склад.xlsx";
                        documentForProject.DocumentType = "Заявка на склад";
                        documentForProject.DocumentPath = savePath + "\\" + documentForProject.DocumentName;
                        // Вношу изменения в базу данных
                        DBCon.DocumentForProjects.Add(documentForProject);
                        DBCon.SaveChanges();
                    }
                }
                // Обновляю грид
                FillDocumentsForProjectGrid();
            }
            catch
            {
                MessageBox.Show("Ошибка при сохранении информации о документе <Заявка на склад> в БД, проверьте ее\nДокумент создан, но не внесен в БД");
            }
        }

        // Настройка колонок грида для первого уровня вложенности
        public void setUpFirstLevelGridColumns()
        {
            dataGridView1.MultiSelect = false;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            dataGridView1.Columns.Clear();
            dataGridView1.ReadOnly = true;
            var column1 = new DataGridViewColumn();
            column1.Name = "Name";
            column1.HeaderText = "Наименование";
            column1.Width = 215;
            column1.CellTemplate = new DataGridViewTextBoxCell();
            var column2 = new DataGridViewColumn();
            column2.Name = "Description";
            column2.HeaderText = "Описание";
            column2.Width = 620;
            column2.CellTemplate = new DataGridViewTextBoxCell();
            var column3 = new DataGridViewColumn();
            column3.Name = "Count";
            column3.HeaderText = "Количество";
            column3.Width = 80;
            column3.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(column1);
            dataGridView1.Columns.Add(column2);
            dataGridView1.Columns.Add(column3);
        }

        // Настройка колонок грида для нулевого уровня вложенности
        public void setUpZeroLevelGridColumns()
        {
            dataGridView1.MultiSelect = false;
            var column1 = new DataGridViewColumn();
            column1.Name = "Article";
            column1.HeaderText = "Изделие";
            column1.Width = 200;
            column1.CellTemplate = new DataGridViewTextBoxCell();
            var column2 = new DataGridViewColumn();
            column2.Name = "Count";
            column2.HeaderText = "Количество компонентов";
            column2.Width = 100;
            column2.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            dataGridView1.Columns.Clear();
            dataGridView1.ReadOnly = true;
            dataGridView1.Columns.Add(column1);
            dataGridView1.Columns.Add(column2);
        }

        // Настройка колонок грида для второго уровня вложенности
        public void setUpSecondLevelGridColumns()
        {
            dataGridView1.MultiSelect = false;
            var column1 = new DataGridViewColumn();
            column1.Name = "Property";
            column1.HeaderText = "Свойство";
            column1.Width = 200;
            column1.CellTemplate = new DataGridViewTextBoxCell();
            var column2 = new DataGridViewColumn();
            column2.Name = "Value";
            column2.HeaderText = "Значение";
            column2.Width = 700;
            column2.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.Columns.Clear();
            dataGridView1.ReadOnly = true;
            dataGridView1.Columns.Add(column1);
            dataGridView1.Columns.Add(column2);
        }

        // Настройка колонок грида для документов по проекту
        public void setUpProjectDocumentsGridColumns()
        {
            dataGridView1.MultiSelect = false;
            dataGridView1.Columns.Clear();
            dataGridView1.Refresh();
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.ReadOnly = true;

            var column1 = new DataGridViewColumn();
            column1.Name = "DocumentName";
            column1.HeaderText = "Имя документа";
            column1.Width = 450;
            column1.CellTemplate = new DataGridViewTextBoxCell();

            var column2 = new DataGridViewColumn();
            column2.Name = "DocumentType";
            column2.HeaderText = "Тип документа";
            column2.Width = 300;
            column2.CellTemplate = new DataGridViewTextBoxCell();

            var column3 = new DataGridViewButtonColumn();
            column3.Name = "Button";
            column3.Width = 100;
            column3.HeaderText = "Действие";
            column3.Text = "Открыть";

            // Скрытый путь к документу
            var column4 = new DataGridViewColumn();
            column4.Name = "DocumentPath";
            column4.CellTemplate = new DataGridViewTextBoxCell();
            column4.Visible = false;

            // Скрытый ID документа
            var column5 = new DataGridViewColumn();
            column5.Name = "DocumentID";
            column5.CellTemplate = new DataGridViewTextBoxCell();
            column5.Visible = false;

            dataGridView1.Columns.Add(column1);
            dataGridView1.Columns.Add(column2);
            dataGridView1.Columns.Add(column3);
            dataGridView1.Columns.Add(column4);
            dataGridView1.Columns.Add(column5);
        }

        // Добавить документ
        private void button6_Click(object sender, EventArgs e)
        {
            this.Hide();
            addDocument.selectedProject = selectedProject;
            addDocument.projectDocDirectoryPath = documentsPath + selectedProject;
            addDocument.Show();
        }

        // Удалить документ
        private void button7_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count < 1)
            {
                MessageBox.Show("Выберите документ для удаления");
            }
            else
            {
                // Номер колонки, где спрятан код документа (отчет с 0)
                const int docIDinGridColumnIndex = 4;
                // Индекс выбранной строки
                int selectedRowIndex = dataGridView1.CurrentRow.Index;
                // Код выбранного документа
                int selectedDocID = Convert.ToInt32(dataGridView1[docIDinGridColumnIndex, selectedRowIndex].Value);
                DocumentForProject documentForProject = new DocumentForProject();
                try
                {
                    // Подключаюсь к БД
                    using (DBContext DBCon = new DBContext())
                    {
                        // Нашел выбранный документ
                        documentForProject = DBCon.DocumentForProjects.Where(
                            o => o.DocumentID == selectedDocID).FirstOrDefault();
                        // Удалил его из БД
                        DBCon.DocumentForProjects.Remove(documentForProject);
                        DBCon.SaveChanges();
                        // Удалил его физически из директории
                        FileInfo fileInfo = new FileInfo(documentForProject.DocumentPath);
                        if (fileInfo.Exists == true)
                        {
                            fileInfo.Delete();
                        }
                        else
                        {
                            MessageBox.Show("Файл на физическом носителе отсутствует.\nПроизводится удаление из базы данных");
                        }
                    }
                    // Обновил грид
                    FillDocumentsForProjectGrid();
                }
                catch
                {
                    MessageBox.Show("Ошибка в БД, проверьте её (функция удаления документа)");
                }
            }
        }

        // Обработка нажатия кнопки "открыть"
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // 0 1 2, во второй колонке с нуля кнопка
            if (e.ColumnIndex == 2)
            {
                string path = dataGridView1[e.ColumnIndex + 1, e.RowIndex].Value.ToString();
                try
                {
                    Process.Start(path);
                }
                catch
                {
                    Exception runErr = new Exception("Ошибка при запуске файла. Проверьте:\n1. Доступ к каталогу\n2. Наличие файла\n3. Наличие ПО для открытия файла\n(Функция обработки кнопки <Открыть>)");
                    throw runErr;
                }
            }
        }

        // Функция заполнения грида документами по проекту
        public void FillDocumentsForProjectGrid()
        {
            try
            {
                dataGridView1.Rows.Clear();
                using (DBContext DBCon = new DBContext())
                {
                    int selectedProjectID = Convert.ToInt32(selectedProject);
                    DocumentForProject[] documentForProjects = DBCon.DocumentForProjects.Where(
                        o => o.ProjectID == selectedProjectID).ToArray();
                    foreach (DocumentForProject documentForProject in documentForProjects)
                    {
                        int docID = documentForProject.DocumentID;
                        string docName = documentForProject.DocumentName;
                        string docType = documentForProject.DocumentType;
                        string docPath = documentForProject.DocumentPath;
                        dataGridView1.Rows.Add(docName, docType, "Открыть", docPath, docID);
                    }
                }
                dataGridView1.ClearSelection();
            }
            catch
            {
                MessageBox.Show("Ошибка заполнения грида с документами по проекту.\nОшибка в функции FillDocumentsForProjectGrid\nПроверьте БД");
            }        
        }
    }
}