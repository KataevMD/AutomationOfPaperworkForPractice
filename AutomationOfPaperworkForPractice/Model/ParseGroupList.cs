using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Table = Microsoft.Office.Interop.Word.Table;
using System.Text.RegularExpressions;
using NameCaseLib;

namespace AutomationOfPaperworkForPractice.Model
{
    static class ParseGroupList
    {
        private static List<Entity.Student> Students = new List<Entity.Student>();
        public static List<Entity.Student> GetStudents() { return Students; }
        public static void ClearStudents() { Students.Clear(); }


        public static void ReadDataFromDocument(string[] PathDocument)
        {
            if (Students.Count > 0)
            {
                Students.Clear();
            }
            Microsoft.Office.Interop.Word.Application wordApp = null;
            Document wordDoc;
            Table wordTable;
            foreach (var path in PathDocument)
            {
                try
                {

                    wordApp = new Microsoft.Office.Interop.Word.Application();

                    object missing = Type.Missing;
                    object fileName = path;

                    wordDoc = wordApp.Documents.Open(ref fileName, ref missing, ref missing, ref missing);

                    wordTable = wordDoc.Tables[1]; //Обращение к таблице результатов студентов за экзамен
                    string groupName = wordTable.Cell(1, 1).Range.Text.Replace("\a", String.Empty).Trim();
                    MessageBox.Show(groupName);
                    int countRowTable = wordTable.Rows.Count;

                    for (int i = 3; i <= countRowTable; i++)
                    {
                        string fullNameStudent = wordTable.Cell(i, 2).Range.Text.Replace("\a", String.Empty).Trim();
                        if (fullNameStudent.Length == 0)
                            break;

                        Entity.Student student = new Entity.Student
                        {
                            FullNameStudent = fullNameStudent,
                            Group = groupName
                        };
                        Students.Add(student);
                    }

                }
                catch (Exception ex)
                {
                    wordApp.ActiveDocument.Close(); //закрытие активного документа
                    wordApp?.Quit();
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    wordApp.ActiveDocument.Close();
                    wordApp?.Quit();
                }
            }
        }
        public static void SetPadej()
        {
            Ru ru = new Ru();
            var listStudent = GetStudents();
            foreach (Entity.Student student in listStudent)
            {
                string[] st = ru.Q(student.FullNameStudent);
                student.FullNameStudent = st[1];
            }
        }

    }
}
