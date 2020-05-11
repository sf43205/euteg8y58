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

        public static void ShowDataInNewApp()
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
                range = range.get_Resize(5, 5);

                if (true)
                {
                    //Create an array.
                    double[,] saRet = new double[5, 5];

                    //Fill the array.
                    for (long iRow = 0; iRow < 5; iRow++)
                    {
                        for (long iCol = 0; iCol < 5; iCol++)
                        {
                            //Put a counter in the cell.
                            saRet[iRow, iCol] = iRow * iCol;
                        }
                    }

                    //Set the range value to the array.
                    range.set_Value(Missing.Value, saRet);
                }

                else
                {
                    //Create an array.
                    string[,] saRet = new string[5, 5];

                    //Fill the array.
                    for (long iRow = 0; iRow < 5; iRow++)
                    {
                        for (long iCol = 0; iCol < 5; iCol++)
                        {
                            //Put the row and column address in the cell.
                            saRet[iRow, iCol] = iRow.ToString() + "|" + iCol.ToString();
                        }
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
