using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;

namespace TrlConsCs
{
    public partial class TableResp
    {
        // Количество блоков превышено ?
        //public const int iAllUnits = 2; // максимальное к-во обрабатываемых блоков
        public const int iAllUnits = 1000; // максимальное к-во обрабатываемых блоков

        public static double[,,] Xi = new double[2, iDep * (iTeams + 1), iAllUnits + 5];
        // 2-Estimations, Points (лист отчета)
        // 35 столбцов таблицы (Total по Department + макс. к-во пар Department/Team)
        // 1000 строки All_Units + 1(Total)
        // Total - 0 строка, блоки - начинать с 1-й

        public static string[] Departments = { "Technical Solution", "Development", "Debugging", "Commissioning", "Documentation" };
        public const int iDep = 5;
        public static int iCurrentDepartment = -1; // текущая стадия (0-4)

        public static string[] Teams = { "Electronics Team", "Firmware Team", "Remote Team", "Mechanics Team", "Commissioning Team", "Software Team" };
        public const int iTeams = 6;
        public static int iCurrentTeam = -1; // текущая команда (0-5)

        public static string[] All_Units = new string[iAllUnits + 5]; //0 - сумма по столбцу, 1-1000 - сумма по блоку
        public static int iAll = 0;  // всего блоков обнаружено
        public static int iCurrentUnit = -1; // текущий блок

        public static string currentCardName;
        public static string currentShortUrl;

        public static int iTotal;
        public static int iCurrTD;
        public static int currentUnit;

        public static double currentEstimate;  // текущее оценочное значение
        public static double currentPoint;  // текущее реальное значение

        public static double minimumSize;

        public static int iError = 0;
        public static int iStartError = 0;

        public static int badges;
        public static int cardRole;
        public static int iLabelsEM;

        public static int iNone;

        public static int[] strErrNumber = new int[iAllUnits + 5];
        public static string[] strErrMessage = new string[iAllUnits + 5];
        public static string[] strErrCardURL = new string[iAllUnits + 5];

        public static ExcelPackage Excel_result { get; set; }
        public static int CurrUnit { get; set; }
        public static double Curr_Estim { get; set; }
        public static double Curr_Point { get; set; }
        public static double MinimumSize { get; set; }

        public static OfficeOpenXml.ExcelPackage excel_result;

        public static void FillExcelSheets(int WorkSheetN)
        {
            for (int WorkSheet = 0; WorkSheet < WorkSheetN - 1; WorkSheet++)
            {
                // Лист для записи оценочных значений WorkSheet = 0
                // Лист для записи реальных значений WorkSheet = 1
                // Лист для записи ошибок WorkSheet = WorkSheetN - 1
                // Лист для записи корректных карточек WorkSheet = WorkSheetN
                OfficeOpenXml.ExcelWorksheet estWorksheet = Excel_result.Workbook.Worksheets[WorkSheet];
                try { FillExcelSheet(WorkSheet, estWorksheet); }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    Console.WriteLine("Press any key");
                    Console.ReadKey();
                    return;
                }
                estWorksheet.Cells[estWorksheet.Dimension.Address].AutoFitColumns(MinimumSize);
            }

            OfficeOpenXml.ExcelWorksheet errWorksheet = Excel_result.Workbook.Worksheets[WorkSheetN - 1];

            for (int i = 0; i < iError; i++) { errWorksheet.Cells[i + 3, 1].Value = strErrNumber[i]; }
            for (int i = 0; i < iError; i++) { errWorksheet.Cells[i + 3, 2].Value = strErrMessage[i]; }
            for (int i = 0; i < iError; i++) { errWorksheet.Cells[i + 3, 3].Value = strErrCardURL[i]; }

            OfficeOpenXml.ExcelWorksheet valueWorksheet = Excel_result.Workbook.Worksheets[WorkSheetN];

            for (int i = 0; i < Trl.iCorrectParse; i++) { valueWorksheet.Cells[i + 2, 1].Value = i + 1; }
            for (int i = 0; i < Trl.iCorrectParse; i++) { valueWorksheet.Cells[i + 2, 2].Value = Trl.strCorrectEstimate[i]; }
            for (int i = 0; i < Trl.iCorrectParse; i++) { valueWorksheet.Cells[i + 2, 3].Value = Trl.strCorrectPoint[i]; }
            for (int i = 0; i < Trl.iCorrectParse; i++) { valueWorksheet.Cells[i + 2, 4].Value = Trl.strCorrectShotrURL[i]; }
        }

        public static void FillExcelSheet(int WorkSheet, OfficeOpenXml.ExcelWorksheet estWorksheet)
        {
            for (int iD = 0; iD < iDep; iD++)
            {
                for (int i = 0; i <= iAll; i++)
                {
                    if (i == iAll)
                        estWorksheet.Cells[i + 4, 1].Value = "Total";
                    else
                        estWorksheet.Cells[i + 4, 1].Value = All_Units[i];
                    for (int j = 0; j <= iTeams; j++)
                    {
                        FillExcelCells(i, j, iD, estWorksheet, WorkSheet);
                    }
                }
            }
        }

        public static void FillExcelCells(int i, int j, int iD, OfficeOpenXml.ExcelWorksheet estWorksheet, int WorkSheet)
        {
            int jD = iD * (iTeams + 1) + j;
            if (i == iAll)
                estWorksheet.Cells[iAll + 4, jD + 2].Formula = "=SUM(" + estWorksheet.Cells[4, jD + 2].Address + ":" + estWorksheet.Cells[iAll + 3, jD + 2].Address + ")";
            else if (j < iTeams)
            {
                estWorksheet.Cells[i + 4, jD + 2].Value = Xi[WorkSheet, jD, i + 1];
                estWorksheet.Cells[i + 4, jD + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                estWorksheet.Cells[i + 4, jD + 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
            }
            ExcelRange range = estWorksheet.Cells[iAll + 5, 1, iAllUnits + 5, jD + 2];
            range.Clear();
        }
    }
}