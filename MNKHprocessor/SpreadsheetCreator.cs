using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Numerics;

namespace MNKHprocessor
{
    class Spreadsheets {
        static bool SAVE = true;

        static int max_dice_probability = 15;
        static double min_percent_resource = 0.1;
        static string spreadsheet_prename_out = "SovietSpreadsheet_T"; //appends {turn number}.xlsx

        public static void RenderSpreadsheet(TurnData turn_data) {
            //https://stackoverflow.com/questions/25134024/clean-up-excel-interop-objects-with-idisposable/25135685#25135685
            //https://stackoverflow.com/questions/17130382/understanding-garbage-collection-in-net/17131389#17131389
            try {
                RenderSpreadsheetInner(turn_data);
            } finally {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        static void RenderSpreadsheetInner(TurnData turn_data) {

            Application xlApp = new Application();
            xlApp.Visible = true;
            Workbook workbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet hidden_calc_sheet = workbook.Worksheets[1];
            hidden_calc_sheet.Name = "Prob_Table";
            setup_probtables(hidden_calc_sheet);
            Worksheet copyable_sheet = workbook.Worksheets.Add();
            Worksheet sheet = workbook.Worksheets.Add();
            sheet.Name = "Main";
            copyable_sheet.Name = "Copyable";
            copyable_sheet.Columns[1].ColumnWidth = 60;
            copyable_sheet.Cells[1, 1] = "[] Plan INSERT NAME HERE";
            copyable_sheet.Cells[2, 1].Formula = string.Concat("=\"-[]\"&Main!F1&\"/", turn_data.max_resources, " Resources (\"&", turn_data.max_resources, "-Main!F1&\" Reserve), \"&Main!D1&\" Dice Rolled\"");


            sheet.Application.ActiveWindow.SplitRow = 2;
            sheet.Application.ActiveWindow.FreezePanes = true;
            int row = 1;
            List<int> categoryRows = new();
            string norolled_actions = "";
            sheet.Cells[1, 1] = "Total";
            sheet.Cells[1, 1].Font.Bold = true;
            row += 1;
            sheet.Rows[2].RowHeight = 30;
            sheet.Rows[2].WrapText = true;
            sheet.Columns[1].ColumnWidth = 40;
            sheet.Cells[2, 1] = "Name";
            sheet.Cells[2, 2] = "Curr Progress";
            sheet.Cells[2, 3] = "Max Progress";
            sheet.Cells[2, 4] = "Dice";
            sheet.Columns[4].ColumnWidth = 4.5;
            sheet.Cells[2, 5] = "RpD";
            sheet.Columns[5].ColumnWidth = 4;
            sheet.Cells[2, 6] = "Cost";
            sheet.Columns[6].ColumnWidth = 5;
            sheet.Cells[2, 7] = "Bonus";
            sheet.Columns[7].ColumnWidth = 5.5;
            sheet.Cells[2, 8] = "Expected Progress";
            sheet.Cells[2, 9] = "Prob Finish";
            sheet.Columns[9].NumberFormat = "0.000%";
            sheet.Cells[2, 10] = "Prob w/ +10 Omake";
            sheet.Columns[10].NumberFormat = "0.000%";
            sheet.Cells[2, 11] = "Prob w/ +15 Omake";
            sheet.Columns[11].NumberFormat = "0.000%";
            sheet.Cells[1, 9] = "Min% resource:";
            sheet.Cells[1, 11] = min_percent_resource;
            sheet.Columns[12].ColumnWidth = 2;
            sheet.Cells[2, 13] = "Roll Required";
            for (int i = 0; i < TurnProcessor.indicator_names.Length; i++) {
                sheet.Cells[2, i + 14] = TurnProcessor.indicator_names[i];
                sheet.Columns[i + 14].ColumnWidth = 10;
            }

            sheet.Columns[1].WrapText = false;
            row += 1;

            foreach (SectionData section in turn_data.sections) {
                row += 1;
                int top_row = row;
                string section_dice_str;
                if (section.forced_count == 0) {
                    section_dice_str = section.dice;
                } else {
                    section_dice_str = string.Concat(section.dice, "+", section.forced_count, " Forced");
                }
                sheet.Cells[top_row, 1] = string.Concat(section.name, " (", section_dice_str, ")");
                sheet.Cells[top_row, 1].Font.Bold = true;
                categoryRows.Add(top_row);
                foreach (ActionData action in section.actions) {
                    //A1=name B2=prog C3=max
                    //D4=dice E5=RpD F6=Cost
                    //G7=Bonus H8=Exp I9=Prob
                    //J10=Omake10 K11=Omake15 L12=XXX
                    //M13=Req N14+=indicators
                    row += 1;
                    sheet.Cells[row, 1] = action.name;
                    if (action.type == ACTION_TYPES.NORMAL) {
                        sheet.Cells[row, 2] = action.prog_curr;
                    }
                    if (action.type == ACTION_TYPES.NORMAL || action.type == ACTION_TYPES.DC) {
                        sheet.Cells[row, 3] = action.prog_max;
                    }
                    sheet.Cells[row, 4] = action.forced_dice;
                    if (action.forced_dice == 0) {
                        sheet.Cells[row, 4].Interior.Color = 0xFFCC99;
                    } else {
                        sheet.Cells[row, 4].Interior.Color = 0x4040F0;
                    }
                    sheet.Cells[row, 5] = action.rpd;
                    sheet.Cells[row, 6].Formula = string.Concat("=D", row, "*E", row);
                    sheet.Cells[row, 6].Interior.Color = 0xB0B0B0;
                    if (action.bonus != 0) {
                        sheet.Cells[row, 7] = action.bonus;
                    }
                    if (action.type == ACTION_TYPES.NORMAL || action.type == ACTION_TYPES.DC) {
                        sheet.Cells[row, 8].Formula = "=IF(D@<=0,\"\",B@+D@*(50.5+G@))".Replace("@", row.ToString());
                        sheet.Cells[row, 13].Formula = "=IF(D@<=0,\"\",C@-B@-(D@*G@))".Replace("@", row.ToString());
                        sheet.Cells[row, 9].Formula = "=IF(D@<=0,\"\",IF(M@<D@,1,IF(M@>D@*100,0,OFFSET(Prob_Table!A1,M@,D@))))".Replace("@", row.ToString());
                        sheet.Cells[row, 10].Formula = "=IF(D@<=0,\"\",IF(M@-10<D@,1,IF(M@-10>D@*100,0,OFFSET(Prob_Table!A1,M@-10,D@))))".Replace("@", row.ToString());
                        sheet.Cells[row, 11].Formula = "=IF(D@<=0,\"\",IF(M@-15<D@,1,IF(M@-15>D@*100,0,OFFSET(Prob_Table!A1,M@-15,D@))))".Replace("@", row.ToString());
                    }
                    if (action.type == ACTION_TYPES.NORMAL) {
                        copyable_sheet.Cells[row, 1].Formula =
                        "=IF(Main!D@<=0,\"\",\"--[]\"&Main!A@& \", \"& Main!D@&\" Dice (\"&Main!F@&\" R), \"&TEXT(Main!I@,\"0%\")&\"/\"&TEXT(Main!K@,\"0%\"))".Replace("@", row.ToString());
                    } else {
                        copyable_sheet.Cells[row, 1].Formula =
                        "=IF(Main!D@<=0,\"\",\"--[]\"&Main!A@& \", \"& Main!D@&\" Dice\")".Replace("@", row.ToString());
                    }
                    for (int i = 0; i < TurnProcessor.indicator_names.Length; i++) {
                        if (action.indicators.ContainsKey(TurnProcessor.indicator_names[i])) {
                            sheet.Cells[row, i + 14].Formula = string.Concat("=IF(OR(D", row, "<=0,K", row, "<=K1),\"\",", action.indicators[TurnProcessor.indicator_names[i]], ")");
                        }
                    }
                    if (action.type == ACTION_TYPES.NOROLL) {
                        norolled_actions = string.Concat(norolled_actions, "-D", row);
                    }
                }
                if (Int32.TryParse(section.dice, out int diceInt)) {
                    sheet.Cells[top_row, 2].Formula = string.Concat("=MAX(0,D", top_row, "-", diceInt+section.forced_count, ")");
                    sheet.Cells[top_row, 2].NumberFormat = "0; -0; ;@";
                    sheet.Cells[top_row, 2].Interior.Color = 0xFFCC99;
                    sheet.Cells[top_row, 3].Formula = "=IF(B@>0,\"Free Dice\",\"\")".Replace("@", top_row.ToString());
                    sheet.Cells[top_row, 3].Interior.Color = 0xFFCC99;
                }
                sheet.Cells[top_row, 4].Formula = string.Concat("=SUM(D", top_row + 1, ":D", row, ")");
                sheet.Cells[top_row, 4].Font.Bold = true;
                sheet.Cells[top_row, 4].Interior.Color = 0xFFCC99;
                sheet.Cells[top_row, 6].Formula = string.Concat("=SUM(F", top_row + 1, ":F", row, ")");
                sheet.Cells[top_row, 6].Font.Bold = true;
                sheet.Cells[top_row, 6].Interior.Color = 0xB0B0B0;
                copyable_sheet.Cells[top_row, 1].Formula = string.Concat(
                    "=\"-[]", section.name, " (\"&Main!D", top_row, "&\"/", section_dice_str, " Dice, \"& Main!F", top_row, "&\" R)\"");
                sheet.get_Range("A" + (top_row + 1), "I" + row).Rows.Group();
                Console.WriteLine(string.Concat(section.name, "(", top_row, "-", row, ")"));
                row += 1;
            }
            sheet.Cells[1, 2].Formula = "=B" + String.Join("+B", categoryRows.ToArray());
            sheet.Cells[1, 3].Formula = string.Concat("/", turn_data.free_dice, " Free");
            sheet.Cells[1, 4].Formula = "=D" + String.Join("+D", categoryRows.ToArray()) + norolled_actions;
            sheet.Cells[1, 6].Formula = "=F" + String.Join("+F", categoryRows.ToArray());
            sheet.Cells[1, 6].Font.Bold = true;
            sheet.Cells[1, 7].Formula = string.Concat("/", turn_data.max_resources, " max");
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            for (int i = 0; i < TurnProcessor.indicator_names.Length; i++) {
                sheet.Cells[1, i + 14].Formula = "=SUMIF(X3:X@,\"<0\")&\" < \"&SUM(X3:X@)&\" < \"&SUMIF(X3:X@,\">0\")".Replace("@", row.ToString()).Replace('X', letters[i + 13]);
            }

            copyable_sheet.Columns[1].AutoFilter(1, "<>");
            if (SAVE) {
                try {
                    workbook.SaveAs(string.Concat(spreadsheet_prename_out, turn_data.turn, ".xlsx"));
                } catch (System.Runtime.InteropServices.COMException) {
                    Console.WriteLine("Save failed!");
                }
            }
            //workbook.Close();
            //https://stackoverflow.com/questions/25134024/clean-up-excel-interop-objects-with-idisposable/25135685#25135685
        }
        //Based on VBA code by SIDAPHutrANipte
        static void setup_probtables(Worksheet s) {
            for (int col = 1; col <= max_dice_probability; col++) {
                Console.WriteLine("Probabilities of " + col + "d100");
                s.Cells[1, col + 1] = string.Concat(col, "d100");
                s.Cells[col + 1, col + 1] = 1M;
                BigInteger POSSIBILITIES = 0;
                for (int row = col + 1; row <= 100 * col; row++) {
                    int dice = col, index = row - 1;
                    for (int k = 0; k <= (index - dice) / 100; k++) {
                        BigInteger temp = nCr(dice, k);
                        temp *= nCr(index - 100 * k - 1, dice - 1);
                        if ((k / 2) * 2 == k) {
                            POSSIBILITIES += temp;
                        } else {
                            POSSIBILITIES -= temp;
                        }
                    }
                    s.Cells[row + 1, col + 1] = 1M - (Decimal)((POSSIBILITIES * 1_000_000) / BigInteger.Pow(100, col)) / 1_000_000M;
                }
            }
        }

        //Based on https://stackoverflow.com/questions/26311984/permutation-and-combination-in-c-sharp
        static BigInteger nCr(int n, int r) {
            // naive: return Factorial(n) / (Factorial(r) * Factorial(n - r));
            return nPr(n, r) / Factorial(r);
        }

        static BigInteger nPr(int n, int r) {
            // naive: return Factorial(n) / Factorial(n - r);
            return FactorialDivision(n, n - r);
        }

        private static BigInteger FactorialDivision(int topFactorial, int divisorFactorial) {
            BigInteger result = 1;
            for (int i = topFactorial; i > divisorFactorial; i--)
                result *= i;
            return result;
        }

        private static BigInteger Factorial(int n) {
            if (n <= 1)
                return 1;
            BigInteger result = 1;
            for (int i = 1; i <= n; i++)
                result *= i;
            return result;
        }
    }
}