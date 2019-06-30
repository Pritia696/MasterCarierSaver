using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportDataToExcel.Interfaces;
using ExportDataToExcel.Models;
using ExportDataToExcel.Services;
using Plugin.Permissions;
using Plugin.Permissions.Abstractions;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Input;
using Xamarin.Forms;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace ExportDataToExcel.ViewModels
{
    public class MainMenuViewModel : BaseViewModel
    {
        List<String> cars = new List<String>
            {
                "CAT 772G 07-88","CAT 772G 21-59","CAT 772G 21-60","CAT 773G 95-04","Volvo A40G 19-51","Volvo A40G 19-52",
                "Volvo A40F 74-01","Volvo A40F 74-02","Volvo A40E 66-29","Volvo A40E 66-30","Volvo A40E 66-31",
                "Volvo A40G 42-75","Volvo A40G 42-76","Volvo A40G 42-77","BelAZ 75-40 650","BelAZ 75-40 61-36"
            };
        public MainMenuViewModel(ReportModel model)
        {
            Title = "Xamarin Developers";
            ExportToExcelCommand = new Command(async () => await ExportDataToExcelAsync(report));
        }

        static string ReplaceHexadecimalSymbols(string txt)
        {
            string r = "[\x00-\x08\x0B\x0C\x0E-\x1F\x26]";
            return Regex.Replace(txt, r, "", RegexOptions.Compiled);
        }

        public async System.Threading.Tasks.Task ExportDataToExcelAsync(ReportModel model)
        {

            // Granted storage permission
            var storageStatus = await CrossPermissions.Current.CheckPermissionStatusAsync(Permission.Storage);

            if (storageStatus != PermissionStatus.Granted)
            {
                var results = await CrossPermissions.Current.RequestPermissionsAsync(new[] { Permission.Storage });
                storageStatus = results[Permission.Storage];
            }

            try
            {
                var path = DependencyService.Get<IExportFilesToLocation>().GetFolderLocation() + "Report" + model.WorkTime.Substring(0, 8) + ".xlsx";
                FilePath = ReplaceHexadecimalSymbols(path);
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();

                    // Задаем колонки и их ширину
                    Columns lstColumns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                    Boolean needToInsertColumns = false;
                    if (lstColumns == null)
                    {
                        lstColumns = new Columns();
                        needToInsertColumns = true;
                    }
                    lstColumns.Append(new Column() { Min = 1, Max = 10, Width = 12, CustomWidth = true });
                    lstColumns.Append(new Column() { Min = 2, Max = 10, Width = 12, CustomWidth = true });
                    lstColumns.Append(new Column() { Min = 3, Max = 10, Width = 12, CustomWidth = true });
                    lstColumns.Append(new Column() { Min = 4, Max = 10, Width = 12, CustomWidth = true });
                    lstColumns.Append(new Column() { Min = 5, Max = 10, Width = 12, CustomWidth = true });
                    lstColumns.Append(new Column() { Min = 6, Max = 10, Width = 12, CustomWidth = true });
                    lstColumns.Append(new Column() { Min = 7, Max = 10, Width = 12, CustomWidth = true });
                    if (needToInsertColumns)
                        worksheetPart.Worksheet.InsertAt(lstColumns, 0);



                    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "MasterReport" };
                    sheets.Append(sheet);

                    workbookPart.Workbook.Save();

                    SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());



                    Row row = new Row { RowIndex = 3 };
                    sheetData.Append(row);


                    InsertCell(model.WorkTime, CellValues.String, 1, row);                                       
                    InsertCell("Мастер:", CellValues.String, 13, row);                    
                    InsertCell(model.MasterName, CellValues.String, 15, row);

                    row = new Row { RowIndex = 5 };
                    sheetData.Append(row);
                    InsertCell("№ а/м", CellValues.String, 1, row);
                    InsertCell("№ погр.", CellValues.String, 1, row);
                    foreach (var t in model.Tecn)
                    {
                        InsertCell(t.Name, CellValues.String, 1, row);
                        InsertCell("", CellValues.String, 1, row);
                    }

                    row = new Row { RowIndex = 6 };
                    sheetData.Append(row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("Водитель", CellValues.String, 1, row);
                    foreach (var t in model.Tecn)
                    {
                        InsertCell(t.DriverName, CellValues.String, 1, row);
                        InsertCell("", CellValues.String, 1, row);

                    }

                    row = new Row { RowIndex = 7 };
                    sheetData.Append(row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("ГП", CellValues.String, 1, row);
                    foreach (var t in model.Tecn)
                    {
                        InsertCell(t.Poroda, CellValues.String, 1, row);
                        InsertCell("", CellValues.String, 1, row);

                    }

                    row = new Row { RowIndex = 8 };
                    sheetData.Append(row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("Напр.", CellValues.String, 1, row);
                    foreach (var t in model.Tecn)
                    {
                        InsertCell(t.WorkPlace, CellValues.String, 1, row);
                        InsertCell("", CellValues.String, 1, row);

                    }
                    UInt32Value i = 9;

                    foreach (var car in cars)
                    {
                        var listmash = new List<Mashine>();
                        var lm = 0;
                        row = new Row { RowIndex = i };
                        sheetData.Append(row);
                        InsertCell(car, CellValues.String, 1, row);
                        foreach (var tex in model.Tecn)
                        {
                            if (tex.Mashines != null)
                            {
                                var mashin = tex.Mashines.Where(x => x.Name == car).FirstOrDefault();
                                if (mashin != null)
                                {
                                    listmash.Add(mashin);
                                }
                            }

                        }
                        if (listmash.Count != 0)
                        {
                            var fmash = listmash.First();
                            InsertCell(fmash.DriverMName, CellValues.String, 1, row);
                            var texmins = listmash.Select(x => x.TechMins.First()).ToList();

                            foreach (var tex2 in texmins)
                            {

                                for (var rrr = 1; rrr <= model.Tecn.Count(); rrr++)
                                {
                                    if (rrr + lm <= tex2.Index)
                                    {
                                        if (rrr + lm == tex2.Index)
                                        {

                                            var mashnow = listmash.Where(x => x.TechMins.Contains(tex2)).FirstOrDefault();
                                            InsertCell(mashnow.Reis, CellValues.String, 1, row);
                                            InsertCell(mashnow.Plecho, CellValues.String, 1, row);
                                            lm++;
                                        }
                                        else
                                        {
                                            InsertCell("", CellValues.String, 1, row);
                                            InsertCell("", CellValues.String, 1, row);
                                        }
                                    }

                                }
                            }
                        }

                        i++;

                    }
                    //комментарий 
                    row = new Row { RowIndex = i };
                    sheetData.Append(row);
                    InsertCell("Комментарий", CellValues.String, 1, row);
                    InsertCell("", CellValues.String, 1, row);
                    foreach (var texnic in model.Tecn)
                    {
                        if (!String.IsNullOrEmpty(texnic.Comment))
                        {
                            InsertCell(texnic.Comment, CellValues.String, 1, row);
                            InsertCell("", CellValues.String, 1, row);
                        }
                        else
                        {
                            InsertCell("", CellValues.String, 1, row);
                            InsertCell("", CellValues.String, 1, row);
                        }

                    }
                    row = new Row { RowIndex = i };
                    sheetData.Append(row);
                    i++;
                    row = new Row { RowIndex = i };
                    sheetData.Append(row);
                    i++;
                    InsertCell("Техника", CellValues.String, 1, row);
                    InsertCell("Водитель", CellValues.String, 1, row);
                    InsertCell("Выполняемая работа", CellValues.String, 1, row);
                    InsertCell("Комментарий", CellValues.String, 1, row);
                    foreach (var texnicplus in model.TecnPlus)
                    {
                        row = new Row { RowIndex = i };
                        sheetData.Append(row);
                        InsertCell(texnicplus.Name, CellValues.String, 1, row);
                        InsertCell(texnicplus.DriverName, CellValues.String, 1, row);
                        InsertCell(texnicplus.Work, CellValues.String, 1, row);
                        InsertCell(texnicplus.Comment, CellValues.String, 1, row);
                        i++;
                    }

                        worksheetPart.Worksheet.Save();
                    MessagingCenter.Send(this, "DataExportedSuccessfully");

                }

            }


            catch (Exception e)
            {
                Debug.WriteLine("ERROR: " + e.Message);
            }

        }

        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }
        /* To create cell in Excel */
        private void InsertCell(string value, CellValues dataType, int cell_num, Row row_index)
        {
            var t = row_index.RowIndex.ToString() + ":" + cell_num.ToString();
            Cell refCell = null;
            var newCell = new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                CellReference = t
            };
            row_index.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue(value);
            newCell.DataType = new EnumValue<CellValues>(dataType);
        }


        public ICommand ExportToExcelCommand { get; set; }

        private ReportModel report;
        public ReportModel Report
        {
            get { return report; }
            set { SetProperty(ref report, value); }
        }

        private string _filePath;
        public string FilePath
        {
            get { return _filePath; }
            set { SetProperty(ref _filePath, value); }
        }

    }
}
