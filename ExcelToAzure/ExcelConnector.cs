using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelToAzure
{
    public static class Xls
    {
        public static void GetArrayFromFile(string filename)
        {
            Sheets objSheets;
            _Worksheet objSheet;
            Range range;

            var xlApp = new Excel.Application();
            var objBook = xlApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);


            try
            {
                try
                {
                    //Get a reference to the first sheet of the workbook.
                    objSheets = objBook.Worksheets;
                    objSheet = (_Worksheet)objSheets.get_Item(1);
                }

                catch (Exception theException)
                {

                    MessageBox.Show(theException.Message, "Missing Workbook?");

                    //You can't automate Excel if you can't find the data you created, so 
                    //leave the subroutine.
                    return;
                }

                //Get a range of data.
                //range = objSheet.get_Range("A1", "E5");
                range = objSheet.UsedRange;
                //Retrieve the data from the range.
                Object[,] saRet;
                saRet = (System.Object[,])range.get_Value(Missing.Value);

                //Determine the dimensions of the array.
                long iRows;
                long iCols;
                iRows = saRet.GetUpperBound(0);
                iCols = saRet.GetUpperBound(1);

                //Build a string that contains the data of the array.
                String valueString;
                valueString = "Array Data\n";

                for (long rowCounter = 1; rowCounter <= iRows && rowCounter < 6; rowCounter++)
                {
                    for (long colCounter = 1; colCounter <= iCols; colCounter++)
                    {

                        //Write the next value into the string.
                        valueString = String.Concat(valueString, (saRet[rowCounter, colCounter] ?? "").ToString() + ", ");
                    }

                    //Write in a new line.
                    valueString = String.Concat(valueString, "\n");
                }

                //Report the value of the array.
                MessageBox.Show(valueString + string.Format("Total number of rows is {0}", iRows.ToString()), "Array Values");
            }

            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        public static void ShowDataInNewApp(List<Record> records)
        {
            Workbooks objBooks;
            Sheets objSheets;
            _Worksheet objSheet;
            Range range;

            try
            {
                // Instantiate Excel and start a new workbook.
                var objApp = new Microsoft.Office.Interop.Excel.Application();
                objBooks = objApp.Workbooks;
                var objBook = objBooks.Add(Missing.Value);
                objSheets = objBook.Worksheets;
                objSheet = (_Worksheet)objSheets.get_Item(1);

                //Get the range where the starting cell has the address
                //m_sStartingCell and its dimensions are m_iNumRows x m_iNumCols.
                range = objSheet.get_Range("A1", Missing.Value);
                range = range.get_Resize(records.Count() + 1, 23);

                if (true)
                {
                    //Create an array.
                    object [,] saRet = new object [records.Count() + 1, 23];
                    //Header
                    saRet[0, 0] = "project.name";
                    saRet[0, 1] = "phase";
                    saRet[0, 2] = "location.code";
                    saRet[0, 3] = "location.name";
                    saRet[0, 4] = "location.bsf";
                    saRet[0, 5] = "level1";
                    saRet[0, 6] = "name1";
                    saRet[0, 7] = "level2";
                    saRet[0, 8] = "name2";
                    saRet[0, 9] = "level3";
                    saRet[0, 10] = "name3";
                    saRet[0, 11] = "level4";
                    saRet[0, 12] = "name4";
                    saRet[0, 13] = "template.code";
                    saRet[0, 14] = "description";
                    saRet[0, 15] = "qty";
                    saRet[0, 16] = "ut";
                    saRet[0, 17] = "price";
                    saRet[0, 18] = "total";
                    saRet[0, 19] = "comments";
                    saRet[0, 20] = "csi_code";
                    saRet[0, 21] = "trade_code";
                    saRet[0, 22] = "estimate_category";

                    //Fill the array.
                    for (int iRow = 0; iRow < records.Count(); iRow++)
                    {
                        saRet[iRow + 1, 0] = records[iRow].location.project.name;
                        saRet[iRow + 1, 1] = records[iRow].phase.phase;
                        saRet[iRow + 1, 2] = records[iRow].location.code;
                        saRet[iRow + 1, 3] = records[iRow].location.name;
                        saRet[iRow + 1, 4] = records[iRow].location.bsf;
                        saRet[iRow + 1, 5] = records[iRow].template.level.level1;
                        saRet[iRow + 1, 6] = records[iRow].template.level.name1;
                        saRet[iRow + 1, 7] = records[iRow].template.level.level2;
                        saRet[iRow + 1, 8] = records[iRow].template.level.name2;
                        saRet[iRow + 1, 9] = records[iRow].template.level.level3;
                        saRet[iRow + 1, 10] = records[iRow].template.level.name3;
                        saRet[iRow + 1, 11] = records[iRow].template.level.level4;
                        saRet[iRow + 1, 12] = records[iRow].template.level.name4;
                        saRet[iRow + 1, 13] = records[iRow].template.code;
                        saRet[iRow + 1, 14] = records[iRow].template.description;
                        saRet[iRow + 1, 15] = records[iRow].qty;
                        saRet[iRow + 1, 16] = records[iRow].template.ut;
                        saRet[iRow + 1, 17] = records[iRow].price;
                        saRet[iRow + 1, 18] = records[iRow].total;
                        saRet[iRow + 1, 19] = records[iRow].comments;
                        saRet[iRow + 1, 20] = records[iRow].csi_code;
                        saRet[iRow + 1, 21] = records[iRow].trade_code;
                        saRet[iRow + 1, 22] = records[iRow].estimate_category;
                    }

                    //Set the range value to the array.
                    range.set_Value(Missing.Value, saRet);
                }

                //Return control of Excel to the user.
                objApp.Visible = true;
                objApp.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }
    }
}
