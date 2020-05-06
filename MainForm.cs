using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using BotAgent.DataExporter;
using ClosedXML.Excel;
using System.Diagnostics;

namespace ReportMaster
{
    public partial class MainForm : Form
    {
        private string csvFileName = "";
        private string xlsFileName = "ReportDoc.xlsx";
        private string[] columnNames = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB" };

        private List<SourceDataRow> sourceDataTable = new List<SourceDataRow>();
        private List<ReportDataRow> reportDataTable = new List<ReportDataRow>();

        public MainForm()
        {
            InitializeComponent();
        }

        private void btnOpenXLSFile_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(csvFileName))
            {
                MessageBox.Show("Не указан CSV-файл для конвертации отчета", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            CreateReport(csvFileName);
        }

        private void CreateReport(string csvFileName)
        {
            btnCreateXLSFileReport.Enabled = false;
            sourceDataTable.Clear();
            reportDataTable.Clear();

            OpenCSVFile(csvFileName);
            CreateXLSFile(xlsFileName);
        }

        private void OpenCSVFile(string scvFileName)
        {            
            Csv csv = new Csv(); 
            csv.FileOpen(scvFileName); 

            // структура csv файла должна быть строгой: дата, номер, плановое значение, актуальное значение
            // данные начинаются с первого столбца, первой строки таблицы
            // разделители для данных должны быть стандартные: между ячейками ',' между строками '"' (OpenOffice предлагает их по-умолчанию для сохранения файла csv)
            for (int i = 0; i < csv.Rows.Count; ++i)
            {
                var csvRow = csv.Rows[i];
                if (csvRow.Count >= 4)
                {
                    SourceDataRow dataRow = new SourceDataRow();                    
                    if (!string.IsNullOrWhiteSpace(csvRow[0]))
                        dataRow.RegDate = DateTime.Parse(csvRow[0]);
                    if (!string.IsNullOrWhiteSpace(csvRow[1]))
                        dataRow.Number = int.Parse(csvRow[1]);
                    if (!string.IsNullOrWhiteSpace(csvRow[2]))
                        dataRow.PlanValue = int.Parse(csvRow[2]);
                    if (!string.IsNullOrWhiteSpace(csvRow[3]))
                        dataRow.ActualValue = int.Parse(csvRow[3]);                    

                    sourceDataTable.Add(dataRow);
                }
            }
        }

        private void CreateXLSFile(string xlsFileName)
        {
            // инициализируем начальные значения данных для xls-отчета
            foreach (var row in sourceDataTable)
                reportDataTable.Add(
                    new ReportDataRow()
                    {
                         RegDate = row.RegDate,
                         Number = row.Number,
                         PlanValue = row.PlanValue,
                         ActualValue = row.ActualValue
                    });
            // формируем первую строку отчета, наименование дня недели + дата  
            List<string> reportDates = reportDataTable.OrderBy(item => item.RegDate).Select(it => it.GetWeekDay() + "   " + it.RegDate.ToShortDateString()).Distinct().ToList();
            reportDates.Add("Итог недель");

            // первые две строки отчета статические информационные, заполняем вручную 
            List<string> firstRow = new List<string>();
            firstRow.Add("");
            firstRow.Add("");
            firstRow.AddRange(reportDates.Select(it => it).ToArray());

            List<string> secondRow = new List<string>();
            secondRow.Add("");
            secondRow.Add("№");
            for (int sr = 0; sr < reportDates.Count; ++sr)
                secondRow.AddRange(new string[] { "Пл.", "Фак.", "Нев."});
            secondRow.Add("%");
            secondRow.Add("%вып.");

            // группируем значения по датам, затем делаем группировку по номерам позиций (Number)
            var dateRows = reportDataTable.GroupBy(it => it.RegDate).ToDictionary(item => item.Key, item => item.ToList());
            var dataTable = reportDataTable.GroupBy(it => it.Number).ToDictionary(item => item.Key, item => item.OrderBy(it => it.RegDate).ToList());

            // формируем xls-таблицу итогового отчета
            // форматирование границ таблицы было требованием тех. задания
            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Report");
                
                // Записываем первую строку отчета
                int row_index = 1;
                int col_index = 0;
                for (int i = 0; i < firstRow.Count; ++i)
                {
                    if (i > 1)
                    {
                        ws.Cell(row_index, columnNames[col_index]).Value = "'" + firstRow[i];
                        ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Range(columnNames[col_index] + row_index + @":" + columnNames[col_index + 2] + row_index).Row(1).Merge();
                        col_index += 3;
                    }
                    else
                    {
                        ws.Cell(columnNames[col_index] + row_index).Value = firstRow[i];
                        ++col_index;
                    }    
                }

                ws.Cell(1, 2).Style.Border.TopBorder = XLBorderStyleValues.Thick;
                ws.Cell(1, 2).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell(1, 2).Style.Border.RightBorder = XLBorderStyleValues.Thick;
                ws.Cell(2, 2).Style.Border.TopBorder = XLBorderStyleValues.Double;
                ws.Cell(2, 2).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell(2, 2).Style.Border.RightBorder = XLBorderStyleValues.Thick;
                ws.Cells("C1:AB1").Style.Border.TopBorder = XLBorderStyleValues.Thick;
                ws.Cells("C1:AB1").Style.Border.BottomBorder = XLBorderStyleValues.Double;

                ws.Cells("B2:AB2").Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                ws.Cells("AA1").Style.Border.RightBorder = XLBorderStyleValues.Thick;
                ws.Cells("AB1").Style.Border.RightBorder = XLBorderStyleValues.Thick;
                ws.Cells("AA2").Style.Border.RightBorder = XLBorderStyleValues.Thick;
                ws.Cells("AB2").Style.Border.RightBorder = XLBorderStyleValues.Thick;
                ws.Cells("AA1").Style.Border.BottomBorder = XLBorderStyleValues.Double;
                ws.Cells("C2:Y2").Style.Border.RightBorder = XLBorderStyleValues.Double;
                ws.Cell("F1").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("F2").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("I1").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("I2").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("L1").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("L2").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("O1").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("O2").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("R1").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("R2").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("U1").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("U2").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("X1").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("X2").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("AA1").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell("AA2").Style.Border.LeftBorder = XLBorderStyleValues.Thick;

                // Записываем вторую строку отчета
                row_index = 2;
                col_index = 0;
                for (int i = 0; i < secondRow.Count; ++i)
                {
                    ws.Cell(row_index, columnNames[col_index]).Value = "'" + secondRow[i];
                    ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    ++col_index;
                }
                
                int count = 1;
                int curr_row = 3;
                row_index = 3;
                int first_col_index = 0;
                col_index = 0;

                double PlanValueAll = 0;
                double ActualValueAll = 0;
                double UnactedValueAll = 0;

                // бежим по всем значениям и записываем данные по каждой дате
                foreach (var dataRow in dataTable)
                {
                    row_index = curr_row;
                    col_index = first_col_index;

                    // первый столбец - №п/п
                    ws.Cell(row_index, columnNames[col_index]).Value = "'" + count;                      
                    ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    ++count;
                    ++col_index;

                    // номер данных в csv
                    ws.Cell(row_index, columnNames[col_index]).Value = "'" + dataRow.Key;                
                    ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.RightBorder = XLBorderStyleValues.Thick;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                    ++col_index;

                    double PlanValueSum = 0;
                    double ActualValueSum = 0;
                    double UnactedValueSum = 0;

                    // бежим по всем датам у этого ключа (Number) и получаем итоговую сумму данных для каждого значения
                    foreach (var item in dataRow.Value)
                    {
                        ws.Cell(row_index, columnNames[col_index]).Value = "'" + item.PlanValue;
                        ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        ws.Cell(row_index, columnNames[col_index]).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                        ws.Cell(row_index, columnNames[col_index]).Style.Border.RightBorder = XLBorderStyleValues.Double;
                        ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                        PlanValueSum += item.PlanValue;
                        ++col_index;
                        ws.Cell(row_index, columnNames[col_index]).Value = "'" + item.ActualValue;
                        ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        ws.Cell(row_index, columnNames[col_index]).Style.Border.RightBorder = XLBorderStyleValues.Double;
                        ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                        ActualValueSum += item.ActualValue;
                        ++col_index;
                        ws.Cell(row_index, columnNames[col_index]).Value = "'" + item.UnactedValue;
                        ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        ws.Cell(row_index, columnNames[col_index]).Style.Border.RightBorder = XLBorderStyleValues.Double;
                        ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                        ++col_index;
                    }

                    // получаем общие суммы по всем данным за неделю 
                    UnactedValueSum = PlanValueSum - ActualValueSum;
                    ws.Cell(row_index, columnNames[col_index]).Value = "'" + PlanValueSum;
                    ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.RightBorder = XLBorderStyleValues.Double;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                    PlanValueAll += PlanValueSum;
                    ++col_index;
                    ws.Cell(row_index, columnNames[col_index]).Value = "'" + ActualValueSum;
                    ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.RightBorder = XLBorderStyleValues.Double;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                    ActualValueAll += ActualValueSum;
                    ++col_index;
                    ws.Cell(row_index, columnNames[col_index]).Value = "'" + UnactedValueSum;
                    ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.RightBorder = XLBorderStyleValues.Double;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                    UnactedValueAll += UnactedValueSum;
                    ++col_index;

                    ws.Cell(row_index, columnNames[col_index]).Value = "'" + string.Format("{0:#.####}", (UnactedValueSum / PlanValueSum) * 100);
                    ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.RightBorder = XLBorderStyleValues.Thick;

                    ++col_index;
                    double actualPercent = (ActualValueSum / PlanValueSum) * 100;
                    ws.Cell(row_index, columnNames[col_index]).Value = "'" + string.Format("{0:#.#}%", actualPercent == 0 ? "0,0" : actualPercent.ToString("#.#"));
                    ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.RightBorder = XLBorderStyleValues.Thick;

                    ++curr_row;
                }

                row_index = curr_row;
                col_index = first_col_index + 1;

                ws.Cell(curr_row, columnNames[col_index]).Value = "'" + "Итог";
                ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.RightBorder = XLBorderStyleValues.Thick;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.TopBorder = XLBorderStyleValues.Thick;

                ++col_index;

                double PlanValueSumDate = 0;
                double ActualValueSumDate = 0;
                double UnactedValueSumDate = 0;

                // получаем общие данные по всем значениям на каждую дату (поле - Итого) и записываем в таблицу
                foreach (var item in dateRows)
                {
                    PlanValueSumDate = 0;
                    ActualValueSumDate = 0;
                    UnactedValueSumDate = 0;

                    PlanValueSumDate = item.Value.Sum(it => it.PlanValue);
                    ActualValueSumDate = item.Value.Sum(it => it.ActualValue);
                    UnactedValueSumDate = PlanValueSumDate - ActualValueSumDate;

                    ws.Cell(row_index, columnNames[col_index]).Value = "'" + PlanValueSumDate;
                    ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    ws.Cell(row_index, columnNames[col_index]).Style.Font.Bold = true;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.TopBorder = XLBorderStyleValues.Thick;
                    ++col_index;
                    ws.Cell(row_index, columnNames[col_index]).Value = "'" + ActualValueSumDate;
                    ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    ws.Cell(row_index, columnNames[col_index]).Style.Font.Bold = true;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.LeftBorder = XLBorderStyleValues.Double;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.TopBorder = XLBorderStyleValues.Thick;
                    ++col_index;
                    ws.Cell(row_index, columnNames[col_index]).Value = "'" + UnactedValueSumDate;
                    ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    ws.Cell(row_index, columnNames[col_index]).Style.Font.Bold = true;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.LeftBorder = XLBorderStyleValues.Double;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                    ws.Cell(row_index, columnNames[col_index]).Style.Border.TopBorder = XLBorderStyleValues.Thick;
                    ++col_index;
                }

                // записываем в таблицу общие суммы по всем данным за неделю 
                ws.Cell(row_index, columnNames[col_index]).Value = "'" + PlanValueAll;
                ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.TopBorder = XLBorderStyleValues.Thick;
                ++col_index;
                ws.Cell(row_index, columnNames[col_index]).Value = "'" + ActualValueAll;
                ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.LeftBorder = XLBorderStyleValues.Double;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.TopBorder = XLBorderStyleValues.Thick;
                ++col_index;
                ws.Cell(row_index, columnNames[col_index]).Value = "'" + UnactedValueAll;
                ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.LeftBorder = XLBorderStyleValues.Double;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.TopBorder = XLBorderStyleValues.Thick;
                ++col_index;

                ws.Cell(row_index, columnNames[col_index]).Value = "'" + string.Format("{0:#.####}", (UnactedValueAll / PlanValueAll) * 100);
                ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.RightBorder = XLBorderStyleValues.Thick;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.TopBorder = XLBorderStyleValues.Thick;
                ++col_index;
                ws.Cell(row_index, columnNames[col_index]).Value = "'" + string.Format("{0:#.#}%", (ActualValueAll / PlanValueAll) * 100);
                ws.Cell(row_index, columnNames[col_index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.RightBorder = XLBorderStyleValues.Thick;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                ws.Cell(row_index, columnNames[col_index]).Style.Border.TopBorder = XLBorderStyleValues.Thick;

                ws.Rows(1, 2).Style.Font.Bold = true;
                ws.Column(2).Style.Font.Bold = true;
                ws.Row(row_index).Style.Font.Bold = true;

                try
                {
                    if (File.Exists(Environment.CurrentDirectory + @"\" + xlsFileName))
                    {
                        if (MessageBox.Show(xlsFileName + " уже существует. Перезаписать файл?", "Внимание!", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                        {
                            try
                            {
                                File.Delete(Environment.CurrentDirectory + @"\" + xlsFileName);
                            }
                            catch (Exception exc)
                            {
                                btnCreateXLSFileReport.Enabled = true;
                                MessageBox.Show("Ошибка удаления файла: " + exc.Message);
                            }
                        }
                        else
                        {
                            btnCreateXLSFileReport.Enabled = true;
                            return;
                        }
                    }

                    try
                    {
                        wb.SaveAs(xlsFileName);
                        tbCSVFileName.Text = "";
                        csvFileName = ""; ;
                    }
                    catch (Exception exc)
                    {
                        MessageBox.Show("Ошибка сохранения файла: " + exc.Message);
                    }

                    btnCreateXLSFileReport.Enabled = true;
                    Process.Start(xlsFileName);
                }
                catch (Exception exc)
                {
                    MessageBox.Show("Ошибка формирования отчета: " + exc.Message);
                }
            }
        }

        private void btnCSVOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog csvOpen = new OpenFileDialog();
            if (csvOpen.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                tbCSVFileName.Text = csvOpen.SafeFileName;
                csvFileName = csvOpen.FileName;
            }
        }
    }
}
