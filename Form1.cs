using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Order_sheet
{
    public partial class FormBase : Form
    {
        Excel.Application appExcel;
        Excel.Workbook book;
        Excel.Worksheet sheetList;
        Excel.Worksheet sheetListPants;
        Excel.Worksheet sheetBust;
        Excel.Worksheet sheetPants;

        string patternColor;
        Regex patternBust = new Regex(@"^\w+\s\d+", RegexOptions.Compiled);

        Dictionary<string, string> bufferBust = new Dictionary<string, string>
        {
            ["Бюст"] = "",
            ["Цвет"] = "",
            ["Чашка"] = "",
            ["Размер"] = "",
            ["Кол-во"] = ""
        };

        Dictionary<string, string> bufferPants = new Dictionary<string, string>
        {
            ["Трусы"] = "",
            ["Цвет"] = "",            
            ["Размер"] = "",
            ["Кол-во"] = ""
        };



        public FormBase()
        {
            InitializeComponent();
            textBoxInput.Text = "D:\\Мувик\\Бланк_трусы.xlsx";
        }

        //Save path to file
        private void buttonOpen_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    textBoxInput.Text = openFileDialog.FileName;                   
                }
            }
            
        }

        private void buttonCreate_Click(object sender, EventArgs e)
        {
            if (textBoxInput.Text == "")
            {
                MessageBox.Show("Укажите путь к файлу.");
                return;
            }

            //Try open file
            if(openExcelFile(textBoxInput.Text))
            {
                //Create two new sheets "Бюсты_Заказ" and "Труселя_Заказ"
                createSheets();

                //Create order list in sheet "Бюсты_Заказ"                
                createBustTable();

                //Create order list in sheet "Трусы_Заказ"
                createPantsTable();

                //Create headers in tables
                createHeaders();

                book.Save();
                book.Close();
                appExcel.Quit();

                MessageBox.Show("Файл успешно обработан.");

                Close();
            }          

        }

        private Boolean openExcelFile(string filePath)
        {
            //Create new app and open file
            appExcel = new Excel.Application();

            try
            {
                book = appExcel.Workbooks.Open(filePath, 0, false);
                return true;

            }
            catch(System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("Файл не найден");
                return false;
            }          

        }

        //Create new sheets for order in .xlsx file
        private void createSheets()
        {
            //Sheet with order list
            sheetList = appExcel.Worksheets[1];
            sheetListPants = appExcel.Worksheets[2];

            //Save color pattern
            Excel.Range range = sheetList.get_Range("A1");
            patternColor = range.Interior.Color.ToString();

            //Create sheet for bust order and pants order
            //If sheets "Бюсты_Заказ" and "Бюсты_Труселя" exist, clear all cells
            try
            {
                sheetBust = appExcel.Worksheets[3];
                sheetBust.Cells.ClearContents();

                sheetPants = appExcel.Worksheets[4];
                sheetPants.Cells.ClearContents();
            }
            catch(System.Runtime.InteropServices.COMException)
            {
                appExcel.Worksheets.Add(After: sheetListPants);
                sheetBust = appExcel.Worksheets[3];
                sheetBust.Name = "Бюсты_Заказ";

                appExcel.Worksheets.Add(After: sheetBust);
                sheetPants = appExcel.Worksheets[4];
                sheetPants.Name = "Труселя_Заказ";
            }            
        }
               
        private void createBustTable()
        {
            string cellColor = "";
            string bustName = "" ;
            string bustColor = "";
            const UInt16 CUP = 4;
            const UInt16 SIZE = 4;
            var rowNumber = 2;
            MatchCollection matches;

            //Walk in table while dont find empty row            
            for (var i = 5; sheetList.Cells[i, 4].Value != null; i++)
            {
                //Save bust name "Бюст 11111"
                if(sheetList.Cells[i, 1].Value != null)
                {
                    matches = patternBust.Matches(sheetList.Cells[i, 1].Value);
                    if (matches.Count > 0)
                        bustName = matches[0].Value;
                }               

                //Save bust color
                if (sheetList.Cells[i, 3].Value != null)
                    bustColor = sheetList.Cells[i, 3].Value;

                for (var j = 5; j <= 11; j++)
                {  
                    //Save cell color to check
                    cellColor = sheetList.Cells[i, j].Interior.Color.ToString();

                    //If the cell is for input 
                    if (cellColor != patternColor && sheetList.Cells[i, j].Value > 0)
                    {
                        bufferBust["Бюст"] = bustName;
                        bufferBust["Цвет"] = bustColor;
                        bufferBust["Чашка"] = sheetList.Cells[i, CUP].Value.ToString();
                        bufferBust["Размер"] = sheetList.Cells[SIZE, j].Value.ToString();
                        bufferBust["Кол-во"] = sheetList.Cells[i, j].Value.ToString();

                        var col = 1;
                        foreach(string value in bufferBust.Values)
                        {
                            sheetBust.Cells[rowNumber, col] = value;
                            col++;
                        }
                        rowNumber++;
                    }                           
                }
            }            
        }

        private void createPantsTable()
        {
            string cellColor = "";
            string pantsName = "";
            string pantsColor = "";            
            const UInt16 SIZE = 4;
            var rowNumber = 2;
            MatchCollection matches;

            //Walk in table while dont find empty row            
            for (var i = 5; sheetListPants.Cells[i, 3].Value != null; i++)
            {
                //Save bust name "Трусы 11111"
                if (sheetListPants.Cells[i, 1].Value != null)
                {
                    matches = patternBust.Matches(sheetListPants.Cells[i, 1].Value);
                    if (matches.Count > 0)
                        pantsName = matches[0].Value;
                }

                //Save pants color
                pantsColor = sheetListPants.Cells[i, 3].Value;

                for (var j = 4; j <= 12; j++)
                {
                    //Save cell color to check
                    cellColor = sheetListPants.Cells[i, j].Interior.Color.ToString();

                    //If the cell is for input 
                    if (cellColor != patternColor && sheetListPants.Cells[i, j].Value > 0)
                    {
                        bufferPants["Трусы"] = pantsName;
                        bufferPants["Цвет"] = pantsColor;                        
                        bufferPants["Размер"] = sheetListPants.Cells[SIZE, j].Value.ToString();
                        bufferPants["Кол-во"] = sheetListPants.Cells[i, j].Value.ToString();

                        var col = 1;
                        foreach (string value in bufferPants.Values)
                        {
                            sheetPants.Cells[rowNumber, col] = value;
                            col++;
                        }
                        rowNumber++;
                    }
                }
            }

        }

        private void createHeaders()
        {
            //Create header in "Бюсты_Заказ"                
            var col = 1;
            foreach (string header in bufferBust.Keys)
            {
                sheetBust.Cells[1, col].Value = header;
                col++;
            }

            Excel.Range range = sheetBust.get_Range("A1");
            range.ColumnWidth = 15;

            //Create header in "Труселя_Заказ"                
            col = 1;
            foreach (string header in bufferPants.Keys)
            {
                sheetPants.Cells[1, col].Value = header;
                col++;
            }

            range = sheetPants.get_Range("A1");
            range.ColumnWidth = 15;
        }
       
    }
}
