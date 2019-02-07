using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Freight_Rate_Forecast
{
    public partial class ForecaseForm : Form
    {
        Boolean isFileExist = false;
        public List<double> tc_currentMonth = new List<double>();
        public List<double> tc_1Month = new List<double>();
        public List<double> tc_2Month = new List<double>();
        public List<double> tc_3Month = new List<double>();
        public List<double> tc_4Month = new List<double>();
        public List<double> tc_currentQ = new List<double>();
        public List<double> tc_1Q = new List<double>();
        public List<double> tc_2Q = new List<double>();
        public List<double> tc_3Q = new List<double>();
        public List<double> tc_4Q = new List<double>();
        public List<double> tc_1Cal = new List<double>();
        public List<double> tc_2Cal = new List<double>();
        public List<double> tc_3Cal = new List<double>();
        public List<double> tc_4Cal = new List<double>();
        public List<double> tc_5Cal = new List<double>();
        public List<double> ctc_currentMonth = new List<double>();
        public List<double> ctc_1Month = new List<double>();
        public List<double> ctc_2Month = new List<double>();
        public List<double> ctc_1Q = new List<double>();
        public List<double> ctc_2Q = new List<double>();
        public List<double> ctc_3Q = new List<double>();
        public List<double> ctc_1Cal = new List<double>();
        public List<double> ctc_2Cal = new List<double>();
        public List<double> ctc_3Cal = new List<double>();

        public ForecaseForm()
        {
            InitializeComponent();
        }

        private void startProcess_Click(object sender, EventArgs e)
        {
            try
            {
                if (daysToForecast.Text != null && daysToForecast.Text != ""
                    && fileTextBox.Text != null && fileTextBox.Text != "" && isFileExist)
                {
                    statusTextBox.Text = "Loading from Excel...";
                    List<string> loadedStringList = loadAndGetExcelData();
                    statusTextBox.Text = "Forecasting values...";
                    partitionForecastValues(loadedStringList);
                    statusTextBox.Text = "Success";
                }
                else
                {
                    MessageBox.Show("Please mention the number of days to forecast and make sure the excel file is selected using the Browse button and try again!");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Following exception occurred: " + ex);
            }
        }

        public List<string> loadAndGetExcelData()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@fileTextBox.Text);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            List<string> columnStringList = new List<string>();
            int count = 0;

            if(rowCount > 3 && colCount > 2)
            {
                for(int rowIndex=4; rowIndex <= rowCount; rowIndex++)
                {
                    count++;
                    //We are only trying to forecast the values present in the first 26 columns for our POC
                    if(colCount >= 26)
                    {
                        string columnString = String.Empty;
                        for (int colIndex = 3; colIndex <= 26; colIndex++)
                        {
                            if (xlRange.Cells[rowIndex, colIndex] != null && xlRange.Cells[rowIndex, colIndex].Value2 != null 
                                && xlRange.Cells[rowIndex, colIndex].Value2.ToString() != String.Empty)
                            {
                                columnString = columnString != String.Empty ? columnString + "," + xlRange.Cells[rowIndex, colIndex].Value2.ToString()
                                : xlRange.Cells[rowIndex, colIndex].Value2.ToString();
                            }
                        }
                        if (!columnString.StartsWith("0,"))
                        {
                            columnStringList.Add(columnString);
                        }
                    }
                }
            }
            return columnStringList;
        }

        public void partitionForecastValues(List<string> loadedStringList)
        {
            if(loadedStringList != null && loadedStringList.Count() > 0)
            {
                foreach(String item in loadedStringList)
                {
                    var stringArray = item.Split(',');
                    if(stringArray != null && stringArray.Count() == 24 && Convert.ToDouble(stringArray[0]) != 0 )
                    {
                        tc_currentMonth.Add(Convert.ToDouble(stringArray[0]));
                        tc_1Month.Add(Convert.ToDouble(stringArray[1]));
                        tc_2Month.Add(Convert.ToDouble(stringArray[2]));
                        tc_3Month.Add(Convert.ToDouble(stringArray[3]));
                        tc_4Month.Add(Convert.ToDouble(stringArray[4]));
                        tc_currentQ.Add(Convert.ToDouble(stringArray[5]));
                        tc_1Q.Add(Convert.ToDouble(stringArray[6]));
                        tc_2Q.Add(Convert.ToDouble(stringArray[7]));
                        tc_3Q.Add(Convert.ToDouble(stringArray[8]));
                        tc_4Q.Add(Convert.ToDouble(stringArray[9]));
                        tc_1Cal.Add(Convert.ToDouble(stringArray[10]));
                        tc_2Cal.Add(Convert.ToDouble(stringArray[11]));
                        tc_3Cal.Add(Convert.ToDouble(stringArray[12]));
                        tc_4Cal.Add(Convert.ToDouble(stringArray[13]));
                        tc_5Cal.Add(Convert.ToDouble(stringArray[14]));
                        ctc_currentMonth.Add(Convert.ToDouble(stringArray[15]));
                        ctc_1Month.Add(Convert.ToDouble(stringArray[16]));
                        ctc_2Month.Add(Convert.ToDouble(stringArray[17]));
                        ctc_1Q.Add(Convert.ToDouble(stringArray[18]));
                        ctc_2Q.Add(Convert.ToDouble(stringArray[19]));
                        ctc_3Q.Add(Convert.ToDouble(stringArray[20]));
                        ctc_1Cal.Add(Convert.ToDouble(stringArray[21]));
                        ctc_2Cal.Add(Convert.ToDouble(stringArray[22]));
                        ctc_3Cal.Add(Convert.ToDouble(stringArray[23]));
                    }
                }

                currMonthTextbox.Text = forecast(tc_currentMonth, Convert.ToInt32(daysToForecast.Text));
                tc1MonTextBox.Text = forecast(tc_1Month, Convert.ToInt32(daysToForecast.Text));
                tc2MonTextBox.Text = forecast(tc_2Month, Convert.ToInt32(daysToForecast.Text));
                tc3MonTextBox.Text = forecast(tc_3Month, Convert.ToInt32(daysToForecast.Text));
                tc4MonTextBox.Text = forecast(tc_4Month, Convert.ToInt32(daysToForecast.Text));
                currQTextBox.Text = forecast(tc_currentQ, Convert.ToInt32(daysToForecast.Text));
                tc2CalTextBox.Text = forecast(tc_1Q, Convert.ToInt32(daysToForecast.Text));
                tc1CalTextBox.Text = forecast(tc_2Q, Convert.ToInt32(daysToForecast.Text));
                curr4QTextBox.Text = forecast(tc_3Q, Convert.ToInt32(daysToForecast.Text));
                curr3QTextBox.Text = forecast(tc_4Q, Convert.ToInt32(daysToForecast.Text));
                curr2QTextBox.Text = forecast(tc_1Cal, Convert.ToInt32(daysToForecast.Text));
                curr1QTextBox.Text = forecast(tc_2Cal, Convert.ToInt32(daysToForecast.Text));
                ctc3CalTextBox.Text = forecast(tc_3Cal, Convert.ToInt32(daysToForecast.Text));
                ctc2CalTextBox.Text = forecast(tc_4Cal, Convert.ToInt32(daysToForecast.Text));
                ctc1CalTextBox.Text = forecast(tc_5Cal, Convert.ToInt32(daysToForecast.Text));
                ctc3QTextBox.Text = forecast(ctc_currentMonth, Convert.ToInt32(daysToForecast.Text));
                ctc2QTextBox.Text = forecast(ctc_1Month, Convert.ToInt32(daysToForecast.Text));
                ctc1QTextBox.Text = forecast(ctc_2Month, Convert.ToInt32(daysToForecast.Text));
                ctc2MonTextBox.Text = forecast(ctc_1Q, Convert.ToInt32(daysToForecast.Text));
                ctc1MonTextBox.Text = forecast(ctc_2Q, Convert.ToInt32(daysToForecast.Text));
                ctcCurMonTextBox.Text = forecast(ctc_3Q, Convert.ToInt32(daysToForecast.Text));
                tc5CalTextBox.Text = forecast(ctc_1Cal, Convert.ToInt32(daysToForecast.Text));
                tc4CalTextBox.Text = forecast(ctc_2Cal, Convert.ToInt32(daysToForecast.Text));
                tc3CalTextBox.Text = forecast(ctc_3Cal, Convert.ToInt32(daysToForecast.Text));
            }
        }

        public string forecast(List<double> valueList, int forecastRange)
        {
            List<double> x_ValuesList = new List<double>(); 
            for (int i = 1; i <= valueList.Count(); i++)
            {
                x_ValuesList.Add(i);
            };
            double[] x_Values = x_ValuesList.ToArray();
            double[] y_Values = valueList.ToArray();
            double x_Avg = 0f;
            double y_Avg = 0f;

            double forecast = 0f;
            double b = 0f;
            double a = 0f;
            double X = 0f; // Forecast

            double tempTop = 0f;
            double tempBottom = 0f;

            // X
            for (int i = 01; i < x_Values.Length; i++)
            {
                x_Avg += x_Values[i];
            }
            x_Avg /= x_Values.Length;

            // Y
            for (int i = 0; i < y_Values.Length; i++)
            {
                y_Avg += y_Values[i];
            }
            y_Avg /= y_Values.Length;

            for (int i = 0; i < y_Values.Length; i++)
            {
                tempTop += (x_Values[i] - x_Avg) * (y_Values[i] - y_Avg);
                tempBottom += Math.Pow(((x_Values[i] - x_Avg)), 2f);
            }


            b = tempTop / tempBottom;
            a = y_Avg - b * x_Avg;

            X = valueList.Count() + forecastRange;
            forecast = a + b * X;

            return forecast.ToString();
        }

        private void browseButton_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                isFileExist = openFileDialog1.CheckFileExists;
                fileTextBox.Text = file;
            }
        }
    }
}
