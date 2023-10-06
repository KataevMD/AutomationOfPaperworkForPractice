using AutomationOfPaperworkForPractice.Model;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AutomationOfPaperworkForPractice
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static string[] massPathToFileListGroup { get; set; }
        public static string[] massPathToFilePattern { get; set; }

        public static OpenFileDialog openFileDialog { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            CreateOpenFileWordDialog();
            ListsPattern.Text += "Выберите файл(ы) которые необходимо заполнить на практику. \n\n";
            PathToFiles.Text += "Выберите файл(ы) с списком группы, для которой необходимо заполнить документы на практику. \n\n";
        }

        private void SelectListGroup_Click(object sender, RoutedEventArgs e)
        {

            if (openFileDialog.ShowDialog() == true)
            {
                massPathToFileListGroup = openFileDialog.FileNames;
                if (massPathToFileListGroup.Length > 0)
                {
                    PathToFiles.Text += " Выбраны документы с списками групп: \n";
                    int i = 1;
                    foreach (var s in massPathToFileListGroup)
                    {
                        PathToFiles.Text += $"   {i}. {s} \n";
                        i++;
                    }
                    btnClearListFileListGroup.IsEnabled = true;
                    CheckSelectFiles();
                }
            }
        }

        private void SelectFilePattern_Click(object sender, RoutedEventArgs e)
        {

            if (openFileDialog.ShowDialog() == true)
            {
                massPathToFilePattern = openFileDialog.FileNames;
                if (massPathToFilePattern.Length > 0)
                {
                    ListsPattern.Text += $" Выбраны файлы для заполнения: \n";
                    int i = 1;
                    foreach (var s in massPathToFilePattern)
                    {
                        ListsPattern.Text += $"   {i}. {s} \n";
                        i++;
                    }
                    btnClearListFilePattern.IsEnabled = true;
                    CheckSelectFiles();
                }

            }

        }

        private void StartParseDocument_Click(object sender, RoutedEventArgs e)
        {
            ParseGroupList.ReadDataFromDocument(massPathToFileListGroup);
            var listStudent = ParseGroupList.GetStudents();
            Lists.Text += $" Список группы: \n";
            int i = 1;
            foreach (var s in listStudent)
            {
                Lists.Text += $"  {i}. {s.FullNameStudent} группы {s.Group} \n";
                i++;
            }
            Lists.Text += $"Конец считывания документа\n\n";

            ParseGroupList.SetPadej();

            Lists.Text += $" Список группы с измененным падежом:\n";
            i = 0;
            foreach (var s in listStudent)
            {
                Lists.Text += $"  {i}. {s.FullNameStudent} группы {s.Group} \n";
                i++;
            }

            FormationOfDocuments.FormatDocumentStudents(massPathToFilePattern);

        }

        private static void CreateOpenFileWordDialog()
        {
            openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                DefaultExt = ".docx",
                Filter = "Word documents|*.doc;*.docx"
            };
        }

        private void CheckSelectFiles()
        {
            if (massPathToFilePattern != null && massPathToFileListGroup != null)
                if (massPathToFilePattern.Length > 0 && massPathToFileListGroup.Length > 0)
                    btnStartParseDocument.IsEnabled = true;
                    
        }

        private void btnClearListFilePattern_Click(object sender, RoutedEventArgs e)
        {
            ClearArray(massPathToFilePattern, btnClearListFilePattern, ListsPattern);
            btnStartParseDocument.IsEnabled = false;
        }

        private void btnClearListFileListGroup_Click(object sender, RoutedEventArgs e)
        {
            ClearArray(massPathToFileListGroup, btnClearListFileListGroup, PathToFiles);
            btnStartParseDocument.IsEnabled = false;
        }

        private void ClearArray(Array clearArray, Button btnLock, TextBlock txtBlAddClearText)
        {
            Array.Clear(clearArray, 0, clearArray.Length);
            btnLock.IsEnabled = false;
            txtBlAddClearText.Text += "ВНИМАНИЕ: СПИСОК ФАЙЛОВ ОЧИЩЕН!!!";
        }
    }
}
