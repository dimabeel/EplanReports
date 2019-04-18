using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using DataBaseLibrary;
using Component = DataBaseLibrary.Component;

namespace SpecificationOfProject.Client
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();

            // Заполняем дерево при запуске
            FillTreeview();
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
        }

        public void FillTreeview()
        {
            // Очищаем перед заполнением дерево
            treeView1.Nodes.Clear();

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
                    foreach (PArticle pArticle in pArticles)
                    {
                        // Ищу описание по изделию
                        string Descr = DBCon.LocationDescriptions.Where(
                            o => o.LocationDescriptionID == pArticle.LocationDesriptionID).
                            Select(o1 => o1.Description).FirstOrDefault();
                        // Заполняю первый уровень
                        projNode.Nodes.Add(Descr);
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

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            const int minComponentNameLength = 3;
            const int minComponentDescriptionLength = 5;
            FunctionManager FM = new FunctionManager();
            // Если выбрано значение в дереве и оно первого уровня вложенности
            if (treeView1.SelectedNode != null && treeView1.SelectedNode.Level == 1)
            {
                // Настриваю грид и колонки
                dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
                dataGridView1.Columns.Clear();
                dataGridView1.ReadOnly = true;
                string selectedItemParentName = treeView1.SelectedNode.Parent.Text;
                string selectedItem = treeView1.SelectedNode.Text;
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

                using (DBContext DBCon = new DBContext())
                {
                    // Нашел проект к которому принадлежит узел в дереве
                    Proj proj = DBCon.Projs.Where(
                        o => o.Name == selectedItemParentName).FirstOrDefault();
                    // Нашел код описания изделия
                    int selectedItemLocationID = DBCon.LocationDescriptions.Where(
                        o => o.Description == selectedItem).Select(
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

            // Если выбрано значение в дереве и оно нулевого уровня вложенности
            if (treeView1.SelectedNode != null && treeView1.SelectedNode.Level == 0)
            {
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
                string selectedItem = treeView1.SelectedNode.Text;
                dataGridView1.Columns.Add(column1);
                dataGridView1.Columns.Add(column2);

                using (DBContext DBCon = new DBContext())
                {
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

            // Если выбрано значение в дереве и оно второго уровня вложенности
            if (treeView1.SelectedNode != null && treeView1.SelectedNode.Level == 2)
            {
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
                // Ищу выбранный компонент и по нему заполняю грид
                string componentName = treeView1.SelectedNode.Name;
                List<ComponentPropertiesInfo> componentPropertiesInfos = FM.GetPropertiesInfos(componentName);
                foreach(ComponentPropertiesInfo componentPropertiesInfo in componentPropertiesInfos)
                {
                    if ((componentPropertiesInfo.Value != "") &&
                        (componentPropertiesInfo.Value != "0") &&
                        (componentPropertiesInfo.Value != null))
                    {
                        dataGridView1.Rows.Add(componentPropertiesInfo.Property, componentPropertiesInfo.Value);
                    }
                }
            }
        }
    }
}
