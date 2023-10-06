using AutomationOfPaperworkForPractice.Entity;
using Microsoft.Office.Interop.Word;
using NameCaseLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace AutomationOfPaperworkForPractice.Model
{
    static class FormationOfDocuments
    {
        
        public static void FormatDocumentStudents(string[] massPathToFilePattern)
        {
            if (massPathToFilePattern.Length > 0)
            {
                var listStudents = ParseGroupList.GetStudents();
                Microsoft.Office.Interop.Word.Application wordApp = null;
                Document wordDoc;
                Table wordTable;
                foreach ( var student in listStudents )
                {
                    foreach (var pathFile  in massPathToFilePattern )
                    {
                        try
                        {

                            wordApp = new Microsoft.Office.Interop.Word.Application();

                            object missing = Type.Missing;
                            object fileName = pathFile;
                            string fName = Path.GetFileName(pathFile);
                            MessageBox.Show(fName);

                            

                            wordDoc = wordApp.Documents.Open(ref fileName, ref missing, ref missing, ref missing);
                            wordTable = wordDoc.Tables[1];
                            switch (fName)
                            {
                                case "Аттестационный лист.docx":
                                    break;

                                case "График практической подготовки.docx":
                                    break;

                                case "Дневник практической подготовки.docx":
                                    break;
                            }


                            
                            string groupName = wordTable.Cell(1, 1).Range.Text.Replace("\a", String.Empty).Trim();

                            int countRowTable = wordTable.Rows.Count;

                            //for (int i = 3; i <= countRowTable; i++)
                            //{
                            //    string fullNameStudent = wordTable.Cell(i, 2).Range.Text.Replace("\a", String.Empty).Trim();
                            //    if (fullNameStudent.Length == 0)
                            //        break;

                            //    Entity.Student student = new Entity.Student
                            //    {
                            //        FullNameStudent = fullNameStudent,
                            //        Group = groupName
                            //    };
                            //    Students.Add(student);
                            //}

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
            }
        }

    }
}
