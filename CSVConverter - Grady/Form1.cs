using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace CSVConverter___Grady
{
    public partial class Form1 : Form
    {
        List<ItemPlain> ItemsList = new List<ItemPlain>();
        string savePath = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void btnGetFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Open XSLX Master";
            ofd.Multiselect = false;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                ItemsList.Clear();
                string excelPath = ofd.FileName;
                label1.Text = Path.GetFileNameWithoutExtension(excelPath);
                savePath = Path.GetDirectoryName(excelPath);

                using (var package = new ExcelPackage(new FileInfo(excelPath)))
                {
                    var ws = package.Workbook.Worksheets[1];
                    //check to make sure sheet has contents
                    if (ws.Dimension == null)
                    {
                        MessageBox.Show("Document is empty.");
                        return;
                    }

                    System.Data.DataTable dt = new System.Data.DataTable();
                    int totalRows = ws.Dimension.End.Row;
                    int totalCols = ws.Dimension.End.Column;

                    // adding in columns to datatable
                    for (int i = 0; i < totalCols; i++)
                        dt.Columns.Add();

                    for (int rowNum = 14; rowNum <= ws.Dimension.Rows; rowNum++)
                    {
                        DataRow dr = dt.NewRow();
                        foreach (var cell in ws.Cells[rowNum, 1, rowNum, totalCols])
                            dr[cell.Start.Column - 1] = cell.Text;
                        dt.Rows.Add(dr);
                    }

                    for (int rowNum = 6; rowNum < dt.Rows.Count; rowNum++)
                    {
                        ItemPlain temp = new ItemPlain();
                        temp.itemCode = dt.Rows[rowNum][0].ToString();
                        temp.desc = replaceSingleQ(dt.Rows[rowNum][1].ToString());
                        temp.Unit = dt.Rows[rowNum][2].ToString();
                        ItemsList.Add(temp);

                        //Items temp = new Items();
                        //temp.Key = dt.Rows[rowNum][0].ToString();
                        //temp.Desc = dt.Rows[rowNum][3].ToString();
                        //temp.Unit = dt.Rows[rowNum][4].ToString();
                        //temp.MPNum = dt.Rows[rowNum][1].ToString();
                        //temp.techSpec = dt.Rows[rowNum][2].ToString();
                        //temp.Quantity = dt.Rows[rowNum][5].ToString();
                        //temp.C1 = dt.Rows[rowNum][6].ToString().Replace("$","");
                        //temp.C2 = dt.Rows[rowNum][8].ToString().Replace("$", "");
                        //temp.C3 = dt.Rows[rowNum][10].ToString().Replace("$", "");
                        //temp.C4 = dt.Rows[rowNum][12].ToString().Replace("$", "");
                        //temp.C5 = dt.Rows[rowNum][14].ToString().Replace("$", "");
                    }
                }
            }
            else
            {
                MessageBox.Show("OFD didnt fucking work.");
            }
        }

        public class ItemPlain
        {
            public string itemCode { get; set; }
            public string desc { get; set; }
            public string Unit { get; set; }
        }

        public string replaceSingleQ(string badstring)
        {

            if (badstring.Contains("'"))
                badstring.Replace("'", "''");
            if (badstring.Contains("\""))
                badstring = badstring.Replace("\"", "\"\"");
            return badstring;
        }

        //converts single excel to multiple excels 
        private void button2_Click(object sender, EventArgs e)
        {
            HashSet<string> Codes1 = new HashSet<string>();
            foreach (var item in ItemsList)
                Codes1.Add(item.itemCode.Substring(0, 7));

            foreach (var Code in Codes1)
            {
                List<ItemPlain> worksheetItems = ItemsList.Where(a => a.itemCode.Substring(0, 7).Contains(Code)).ToList();
                //HashSet<string> Worksheets = new HashSet<string>();
                //foreach (Items item in Workbook)
                //    Worksheets.Add(item.Key.Substring(5, 4));

                //ensuring a blank workbook is used to store export data 
                FileInfo newFile = new FileInfo(savePath + "\\" + Code + ".xlsx");
                if (newFile.Exists)
                    newFile.Delete();  // ensures we create a new workbook

                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    //foreach (string sheetName in Worksheets)
                    //{
                    System.Data.DataTable dt = new DataTable();
                    dt.Columns.Add("");
                    dt.Columns.Add("");
                    dt.Columns.Add("");

                    // add a new worksheet to the empty workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("1");

                    foreach (ItemPlain item in worksheetItems)
                    {
                        DataRow dr = dt.NewRow();
                        dr[0] = item.itemCode;
                        dr[1] = item.desc;
                        dr[2] = item.Unit;
                        dt.Rows.Add(dr);
                    }
                    //loads datatable containing layers to worksheet 
                    worksheet.Cells.LoadFromDataTable(dt, true);
                    package.Save();
                }
            }
        }

        private void btnMakeCsv_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Open XSLX Master";
            ofd.Multiselect = true;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                foreach (string excelPath in ofd.FileNames)
                {
                    label1.Text = Path.GetFileNameWithoutExtension(excelPath);
                    savePath = Path.GetDirectoryName(excelPath);
                    string folderName = Path.GetFileNameWithoutExtension(excelPath);

                    using (var package = new ExcelPackage(new FileInfo(excelPath)))
                    {
                        if (!Directory.Exists(savePath + "\\" + folderName))
                            System.IO.Directory.CreateDirectory(savePath + "\\" + folderName);
                        int sheetsNumber = package.Workbook.Worksheets.Count;

                        for (int i = 1; i <= sheetsNumber; i++)
                        {
                            string TargetFile = savePath + "\\" + folderName + "\\" + package.Workbook.Worksheets[i].Name + ".csv";
                            if (System.IO.File.Exists(TargetFile))
                                System.IO.File.Delete(TargetFile);
                            EpplusCsvConverter.ConvertToCsv(package, TargetFile, i);
                        }
                    }
                    //move file to folder 
                    System.IO.File.Move(excelPath, (savePath + "\\" + folderName + "\\" + folderName + ".xlsx"));
                }
            }
        }
    }
}

public static class EpplusCsvConverter
{
    public static void ConvertToCsv(this ExcelPackage package, string TargetFile, int sheetNum)
    {
        var worksheet = package.Workbook.Worksheets[sheetNum];

        var maxColumnNumber = worksheet.Dimension.End.Column;
        var currentRow = new List<string>(maxColumnNumber);
        var totalRowCount = worksheet.Dimension.End.Row;
        var currentRowNum = 1;

        using (var writer = new StreamWriter(TargetFile, false, Encoding.UTF8))
        {
            while (currentRowNum <= totalRowCount)
            {
                BuildRow(worksheet, currentRow, currentRowNum, maxColumnNumber);
                WriteRecordToFile(currentRow, writer, currentRowNum, totalRowCount);
                currentRow.Clear();
                currentRowNum++;
            }
        }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="record">List of cell values</param>
    /// <param name="sw">Open Writer to file</param>
    /// <param name="rowNumber">Current row num</param>
    /// <param name="totalRowCount"></param>
    /// <remarks>Avoiding writing final empty line so bulk import processes can work.</remarks>
    private static void WriteRecordToFile(List<string> record, StreamWriter sw, int rowNumber, int totalRowCount)
    {
        var commaDelimitedRecord = record.ToDelimitedString(",");
        if (rowNumber == totalRowCount)
        {
            sw.Write(commaDelimitedRecord);
        }
        else
        {
            sw.WriteLine(commaDelimitedRecord);
        }
    }

    private static void BuildRow(ExcelWorksheet worksheet, List<string> currentRow, int currentRowNum, int maxColumnNumber)
    {
        for (int i = 1; i <= maxColumnNumber; i++)
        {
            var cell = worksheet.Cells[currentRowNum, i];
            if (cell == null)
            {
                // add a cell value for empty cells to keep data aligned.
                AddCellValue(string.Empty, currentRow);
            }
            else
            {
                AddCellValue(GetCellText(cell), currentRow);
            }
        }
    }

    private static string DuplicateTicksForSql(this string s)
    {
        return s;
    }

    /// <summary>
    /// Takes a List collection of string and returns a delimited string.  Note that it's easy to create a huge list that won't turn into a huge string because
    /// the string needs contiguous memory.
    /// </summary>
    /// <param name="list">The input List collection of string objects</param>
    /// <param name="qualifier">
    /// The default delimiter. Using a colon in case the List of string are file names,
    /// since it is an illegal file name character on Windows machines and therefore should not be in the file name anywhere.
    /// </param>
    /// <param name="insertSpaces">Whether to insert a space after each separator</param>
    /// <returns>A delimited string</returns>
    /// <remarks>This was implemented pre-linq</remarks>
    public static string ToDelimitedString(this List<string> list, string delimiter = ":", bool insertSpaces = false, string qualifier = "", bool duplicateTicksForSQL = true)
    {
        var result = new StringBuilder();
        for (int i = 0; i < list.Count; i++)
        {
            string initialStr = duplicateTicksForSQL ? DuplicateTicksForSql(list[i]) : list[i];
            result.Append((qualifier == string.Empty) ? initialStr : string.Format("{1}{0}{1}", initialStr, qualifier));
            if (i < list.Count - 1)
            {
                result.Append(delimiter);
                if (insertSpaces)
                {
                    result.Append(' ');
                }
            }
        }
        return result.ToString();
    }

    /// <summary>
    /// Can't use .Text: http://epplus.codeplex.com/discussions/349696
    /// </summary>
    /// <param name="cell"></param>
    /// <returns></returns>
    private static string GetCellText(ExcelRangeBase cell)
    {
        return cell.Value == null ? string.Empty : cell.Value.ToString();
    }

    private static void AddCellValue(string s, List<string> record)
    {
        record.Add(string.Format("{0}{1}{0}", '"', s));
    }
}

public class Items
{
    public string Key { get; set; }
    public string Desc { get; set; }
    public string Unit { get; set; }
    public string MPNum { get; set; }
    public string techSpec { get; set; }
    public string Quantity { get; set; }
    public string C1 { get; set; }
    public string C2 { get; set; }
    public string C3 { get; set; }
    public string C4 { get; set; }
    public string C5 { get; set; }
}

