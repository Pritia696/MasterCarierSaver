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

        public async System.Threading.Tasks.Task ExportDataToExcelAsync( ReportModel model)
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
                var path = DependencyService.Get<IExportFilesToLocation>().GetFolderLocation() + "Report" + model.WorkTime.Substring(0,8) + ".xlsx";
                FilePath = ReplaceHexadecimalSymbols(path);
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();

                    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "MasterReport" };
                    sheets.Append(sheet);

                    workbookPart.Workbook.Save();

                    SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                          

                    Row row = new Row { RowIndex = 3 };
                    sheetData.Append(row);


                    InsertCell(model.WorkTime, CellValues.String, 1, row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("Мастер:", CellValues.String, 13, row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell(model.MasterName, CellValues.String, 15, row);

                    row = new Row { RowIndex = 5 };
                    sheetData.Append(row);
                    InsertCell("№ а/м", CellValues.String, 1, row);
                    InsertCell("№ погр.", CellValues.String, 1, row);


                    row = new Row { RowIndex = 6 };
                    sheetData.Append(row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("Водитель", CellValues.String, 1, row);

                    row = new Row { RowIndex = 7 };
                    sheetData.Append(row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("ГП", CellValues.String, 1, row);

                    row = new Row { RowIndex = 8 };
                    sheetData.Append(row);
                    InsertCell("", CellValues.String, 1, row);
                    InsertCell("Напр.", CellValues.String, 1, row);

                    // Insert the header row to the Sheet Data
                    var Developer = new Technique
                    {
                        DriverName = "djlntkm1",
                        Name = "123"
                    };
                    var Developer2 = new Technique
                    {
                        DriverName = "djlntkm2",
                        Name = "456"
                    };
                    var Developers = new List<Technique>();
                    Developers.Add(Developer);
                    Developers.Add(Developer2);

                    // Add each product
                    foreach (var d in Developers)
                    {
                        row = new Row { RowIndex = 9 };
                        row.Append(
                            ConstructCell(d.Name.ToString(), CellValues.String),
                            ConstructCell(d.DriverName, CellValues.String)
                            );
                            
                        sheetData.AppendChild(row);
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
