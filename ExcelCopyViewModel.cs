using System;
using System.Collections.Generic;
using System.Windows.Input;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace CopyExcelData
{
    public class ExcelCopyViewModel : NotifyPropertyChanged
    {
        #region PrivateFields

        private string[] _sourceFilePath;

        private string _sourceFileName;
        private string _destinationFilePath;
        private string _destinationFileName;
        private string _message;

        private bool _controlChecked = true;
        private bool _droughtChecked;
        private bool _isCopying;

        private readonly List<string> _sourceWorksheetNames = new List<string> { "FvFm", "YIIef", "ETR", "NPQs", "qP", "qN", "qL", "Chl Idx", "Ari Idx" };

        #endregion

        #region PublicFields

        public string[] SourceFilesPath
        {
            get => _sourceFilePath;
            set
            {
                _sourceFilePath = value;
                OnPropertyChanged();
            }
        }
        public string SourceFileName
        {
            get => _sourceFileName;
            set
            {
                _sourceFileName = value;
                OnPropertyChanged();
            }
        }

        public string DestinationFilePath
        {
            get { return _destinationFilePath; }
            set
            {
                _destinationFilePath = value;
                OnPropertyChanged();
            }
        }

        public string DestinationFileName
        {
            get { return _destinationFileName; }
            set
            {
                _destinationFileName = value;
                OnPropertyChanged();
            }
        }

        public bool ControlChecked
        {
            get => _controlChecked;
            set
            {
                _controlChecked = value;
                OnPropertyChanged();
            }
        }

        public bool DroughtChecked
        {
            get => _droughtChecked;
            set
            {
                _droughtChecked = value;
                OnPropertyChanged();
            }
        }

        public bool IsCopying
        {
            get { return _isCopying; }
            set
            {
                _isCopying = value;
                OnPropertyChanged();
            }
        }

        public string Message
        {
            get { return _message; }
            set
            {
                _message = value;
                OnPropertyChanged();
            }
        }

        #endregion

        private void SelectSourceFiles(object obj)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Title = "Select an Excel File";

            if (openFileDialog.ShowDialog() == true)
            {
                SourceFilesPath = openFileDialog.FileNames;
                SourceFileName = string.Join("; ", openFileDialog.SafeFileNames);
            }
        }

        private void SelectDestinationFile(object obj)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select an Excel File";

            if (openFileDialog.ShowDialog() == true)
            {
                DestinationFilePath = openFileDialog.FileName;
                DestinationFileName = openFileDialog.SafeFileName;
            }
        }

        private void CopyData(object obj)
        {
            _message = "";

            if (SourceFilesIsEmpty())
            {
                Message = "Please select an Excel file first.";
                return;
            }

            IsCopying = true;

            Application sourceExcel = null;
            Application destinationExcel = null;

            Workbook sourceExcelWorkbook = null;
            Workbook desctinationWorkbook = null;

            Worksheet sourceWorksheet = null;

            try
            {
                destinationExcel = new Application();
                desctinationWorkbook = destinationExcel.Workbooks.Open(DestinationFilePath);

                foreach (var sourceFile in SourceFilesPath)
                {
                    sourceExcel = new Application();
                    sourceExcelWorkbook = sourceExcel.Workbooks.Open(sourceFile);
                    sourceWorksheet = sourceExcelWorkbook.ActiveSheet;

                    _sourceWorksheetNames.ForEach(sheetName =>
                    {
                        Copy(sourceWorksheet, desctinationWorkbook, sheetName);
                    });

                    desctinationWorkbook.Save();

                    sourceExcelWorkbook?.Close(false);
                    sourceExcel?.Quit();
                }

                Message = $"Success!";
            }
            catch (Exception ex)
            {
                Message = $"Error: {ex.Message}";
            }
            finally
            {
                desctinationWorkbook?.Close();
                destinationExcel?.Quit();

                IsCopying = false;
            }
        }

        private bool SourceFilesIsEmpty()
        {
            return SourceFilesPath == null || SourceFilesPath.Length == 0;
        }

        private bool TryGetCellByNameSheet(Worksheet worksheet, string nameSheet, out Range findedCell)
        {
            findedCell = null;

            var nameParam = GetParamName(nameSheet);

            findedCell = worksheet.Cells.Find(nameParam, LookAt: XlLookAt.xlWhole);
            if (findedCell == null)
            {
                Message = $"Could not find cell with name {nameParam}.";
                return false;
            }
            return true;
        }

        private string GetParamName(string nameSheet)
        {
            switch (nameSheet)
            {
                case "FvFm":
                    return "Fv/Fm";
                case "YIIef":
                    return "Fq'/Fm'";
                case "ETR":
                    return "rETR";
                case "NPQs":
                    return "NPQ";
                case "qP":
                    return "qP";
                case "qN":
                    return "qN";
                case "qL":
                    return "qL";
                case "Chl Idx":
                    return "ChlIdx";
                case "Ari Idx":
                    return "AriIdx";

                default: return "";
            }
        }

        private void Copy(Worksheet sourceWorksheet, Workbook desctinationWorkbook, string sheetName)
        {
            if (!TryGetCellByNameSheet(sourceWorksheet, sheetName, out var findedCell)) return;

            Worksheet desctinationWorkSheet = (Worksheet)desctinationWorkbook.Sheets[sheetName];

            var firstEmptyCellNumberInColumnCorH = GetEmptyCellNumberFromColumn(desctinationWorkSheet, ControlChecked ? "C1" : "H1");
            var firstEmptyCellNumberInColumnDorI = GetEmptyCellNumberFromColumn(desctinationWorkSheet, ControlChecked ? "D1" : "I1");

            int startRow = findedCell.Row + 2;
            int numRowsForCopy = 6;

            object[,] dataForColumnCorH = new object[3, 1];
            object[,] dataForColumnDorI = new object[3, 1];

            for (int i = 0; i < numRowsForCopy; i++)
            {
                if (i < 3)
                {
                    dataForColumnCorH[i, 0] = sourceWorksheet.Cells[startRow + i, findedCell.Column].Value2;
                    continue;
                }

                dataForColumnDorI[i - 3, 0] = sourceWorksheet.Cells[startRow + i, findedCell.Column].Value2;
            }

            var destinationRangeCorH = GetDestinationRange(desctinationWorkSheet, ControlChecked ? "C" : "H", firstEmptyCellNumberInColumnCorH);
            var destinationRangeDorI = GetDestinationRange(desctinationWorkSheet, ControlChecked ? "D" : "I", firstEmptyCellNumberInColumnDorI);

            destinationRangeCorH.Value2 = dataForColumnCorH;
            destinationRangeDorI.Value2 = dataForColumnDorI;
        }

        private static Range GetDestinationRange(Worksheet desctinationWorkSheet, string columnName, int firstEmptyCellNumber)
        {
            return desctinationWorkSheet.Range[$"{columnName}{firstEmptyCellNumber}", $"{columnName}{firstEmptyCellNumber + 2}"];
        }

        private int GetEmptyCellNumberFromColumn(Worksheet outputWorksheet, string nameColumn)
        {
            var column = outputWorksheet.Range[nameColumn];
            var emptyRowNumberColumn = column.get_End(XlDirection.xlDown).Row;
            return emptyRowNumberColumn + 1;
        }

        public ICommand CopyDataCommand => new UcCommand(CopyData);
        public ICommand SelectSourceFileCommand => new UcCommand(SelectSourceFiles);
        public ICommand SelectDestinationFileCommand => new UcCommand(SelectDestinationFile);
    }
}