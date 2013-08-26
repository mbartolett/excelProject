using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util.Collections;

namespace npoiTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            label1.Visible = false;
            cboSheets.Visible = false;
            lblDescription.Visible = false;
            textBox1.Visible = false;
        }

        HSSFWorkbook workBook;
        DataSet dataSet1 = new DataSet();

        void InitializeWorkbook(string path)
        {
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                 workBook = new HSSFWorkbook(file);
            }
        }

        void ConvertToDataTable()
        {
            //create sheet obj
             ISheet sheet = workBook.GetSheetAt(cboSheets.SelectedIndex);
            //dtRushMuthaFucker(sheet);
            dataSet1.Tables.Add(dtRushMuthaFucker(sheet));
            //return! dtrushmuthafucker does it all! i knew it could be done in one function
            return;

            #region im hiding this shit, cause dtrush ends all fuckers
            getDealColumnsSheetLevel(sheet);

            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            
            //create datatable
            DataTable dt = new DataTable();
            
            //add columns to datatable
            for (int i = 0; i < 32; i++)
            {
                dt.Columns.Add(Convert.ToChar(((int)'A') + i).ToString());
            }

            //load worksheet into datatable
            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                DataRow dr = dt.NewRow();

                for (int j = 0; j < row.LastCellNum; j++)
                {
                    ICell cell = row.GetCell(j);

                    if (cell == null)
                    {
                        dr[j] = null;
                    }
                    else
                    {
                        dr[j] = cell.ToString();
                    }
                }
                dt.Rows.Add(dr);                
            }

            int iCurrentSheetColumns = NumberofDealColumns(dt);
            List<int> lCustomerNumbers = iCustNumbers(dt);

            //create new datatable and copyin wanted columns
            DataTable updatedTable = dt.DefaultView.ToTable(false, "G", "I", "J","K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "V");

            //take out emtpy rows
            updatedTable = updatedTable.Rows.Cast<DataRow>().Where(row => !row.ItemArray.All(field => field is System.DBNull || string.Compare((field as string).Trim(), string.Empty) == 0)).CopyToDataTable();

            //create new table for rows to transfer
            DataTable currentTable = new DataTable();
            
            for (int i = 0; i < 32; i++)
            {
                currentTable.Columns.Add(Convert.ToChar(((int)'A') + i).ToString());
            }
            //determine if the value in column g is an int
            DataColumn col = updatedTable.Columns["G"];
            foreach (DataRow row in updatedTable.Rows)
            {
                object value = row[col];
                
                //check if cell is null
                if (value != DBNull.Value)
                {
                    //check if value is an int
                    //first cast the value as a string
                    string sValue = value.ToString().Trim();
                    //find the index of the last space in the string
                    Int32 iLastSpace = sValue.LastIndexOf(' ');
                    //find the length of the string
                    Int32 iLastChar = sValue.Length;
                    //if neither of those were zero lets move on
                    if (iLastSpace > 0 & iLastChar > 0)
                    {
                        //grab the last part of the string this is where your deal # lives
                        string sLastValue = sValue.Substring(iLastSpace, iLastChar - iLastSpace).Trim();
                        int num;
                        //if it is not an int we delete the row.
                        if (int.TryParse(sLastValue, out num))
                        {
                            row.SetField(col, sLastValue);
                            //this is messy and not the best way but it works
                            for (int j = 8; j < dt.Columns.Count; j++)
                            {
                                string sUpdate;
                                sUpdate = dt.Rows[16][j].ToString().Trim();
                                if (!string.IsNullOrEmpty(sUpdate))
                                {
                                    int iStart;
                                    int iLast;
                                    iStart = sUpdate.LastIndexOf(" ");
                                    iLast = sUpdate.Length;
                                    sUpdate = sUpdate.Substring(iStart, iLast - iStart).Trim();
                                    sUpdate = sUpdate + " " + row[j - 7].ToString();
                                    row.SetField(j - 7, sUpdate);
                                }
                            }
                            currentTable.ImportRow(row);
                            currentTable.AcceptChanges();
                        }
                        else
                        {
                            //was not a valid int, get rid of row.
                            row.Delete();
                            
                        }
                    }
                    else
                    {
                        //was no space in this cell delete the row
                        row.Delete();                        
                    }
                }
                else 
                {
                    //value was null, delete the row
                    row.Delete();
                }
            }
            //have to tell the table to accept changes so all the deleted rows actually delete
            updatedTable.AcceptChanges();        
            //add it to the dataset
            dataSet1.Tables.Add(currentTable);
            #endregion

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btn_import_Click(object sender, EventArgs e)
        {
            //InitializeWorkbook(@"test.xls");
            ConvertToDataTable();            
            //have to autogenerate columns or set them up yourself... aint nobody got time for that
            dataGridView1.AutoGenerateColumns = true;
            //bind return table to datagridview, should show up
            dataGridView1.DataSource = dataSet1.Tables[0];
            
        }

        //pass the full datatable in to here and it will return the number of deal columns
        private int NumberofDealColumns(DataTable dtReadColumns)
        {
            //set the counter above the try so when it throws error we return the last good column
            int iCounter = 0;
            try
            {
                //current dealID we are trying to read
                int iDealNum = 1;
                
                for (int j = 8; iDealNum > 0; j++)
                {
                    //read the column header here
                    string sUpdate;
                    sUpdate = dtReadColumns.Rows[16][j].ToString().Trim();
                    int iStart;
                    int iLast;
                    iStart = sUpdate.LastIndexOf(" ");
                    iLast = sUpdate.Length;
                    //it actually exits here when the substring fails due to negative number
                    sUpdate = sUpdate.Substring(iStart, iLast - iStart).Trim();
                    //if we got a value out of the cell, make sure it is an int and add it to the counter
                    if (int.TryParse(sUpdate, out iDealNum))
                    {
                        iCounter++;
                    }
                    else
                    {
                        iDealNum = -1;
                    }
                }

                //either way, return the counter we know how many deals we have in this sheet.
                return iCounter;
            }
            catch
            {
                return iCounter;
            }
        }

        private List<int> iCustNumbers(DataTable dtToScan)
        {
            List<int> iCustomers = new List<int>();
            try
            {
                for (int j = 19; j < dtToScan.Rows.Count; j++)
                {
                    //read the column header here
                    string sUpdate;
                    sUpdate = dtToScan.Rows[j][6].ToString().Trim();
                    if (!string.IsNullOrEmpty(sUpdate))
                    {
                        int iStart;
                        int iLast;
                        iStart = sUpdate.LastIndexOf(" ");
                        iLast = sUpdate.Length;
                        //it actually exits here when the substring fails due to negative number
                        sUpdate = sUpdate.Substring(iStart, iLast - iStart).Trim();
                        //if we got a value out of the cell, make sure it is an int and add it to the counter
                        int iCustomerNum;
                        if (int.TryParse(sUpdate, out iCustomerNum))
                        {
                            iCustomers.Add(iCustomerNum);
                        }                        
                    }
                }
                return iCustomers;
            }
            catch
            {
                return iCustomers;
            }
        }

        private int getDealColumnsSheetLevel(ISheet sheet)
        {
            int iReturnVal = 0;
            try
            {
                System.Collections.IEnumerator rows = sheet.GetRowEnumerator();               
                IRow test = sheet.GetRow(16);
                iReturnVal = test.LastCellNum - 8;              

                return iReturnVal;
            }
            catch
            {
                return iReturnVal;
            }
        }

        private DataTable dtRushMuthaFucker(ISheet sheet)
        {
            dataSet1.Tables.Clear();
            //create datatable cause we have to fucking return a data table
            DataTable dtMaster = new DataTable("bitchtits");
            try
            {
                //had to make time for this.. fuck
                dtMaster.Columns.Add("CustNumber", typeof(string));
                dtMaster.Columns.Add("DealNumber", typeof(string));
                dtMaster.Columns.Add("Quantity", typeof(int));

                //get those motha fucking rows
                System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
                //this row has the motha fucking deal columns bitch.. save that shit for later use
                IRow DealRows = sheet.GetRow(16);
                //count the number of fucking deals, so we can step through it later
                int iColumnCount = DealRows.LastCellNum;

                //start at row 19, cause thats where the fucking customers start, then go till the last fucking row
                for (int counter = 19; counter < sheet.LastRowNum; counter++)
                {
                    //get current fucking customer row
                    IRow cRow = sheet.GetRow(counter);
                    string sTrim;
                    //gotta try catch that shit, cause those blank rows fuck shit up
                    try
                    {
                        sTrim = cRow.GetCell(6).ToString();
                    }
                    catch
                    {
                        //blank row blows the fuck up set trim to empty and move the fuck on
                        sTrim = string.Empty;
                    }

                    if (!string.IsNullOrEmpty(sTrim))
                    {
                        //gotta get the customer number out of those two numbers, put on your magic hat its time for sorcery
                        int iStart;
                        int iLast;
                        iStart = sTrim.LastIndexOf(" ");
                        iLast = sTrim.Length;
                        
                        sTrim = sTrim.Substring(iStart, iLast - iStart).Trim();
                        //if we got a value out of the cell, make sure it is an int and add it to the counter
                        int iCustomerNum;
                        //int tryparse is fucking awesome! 
                        if (int.TryParse(sTrim, out iCustomerNum))                        
                        {                                
                            //ok wizard, we got the customer number, lets read some fucking values ruler style
                            for (int columnStart = 8; columnStart < iColumnCount; columnStart++)
                            {
                                //again its time for string magic, cause the excel sheet is fucking stupid.. get the deal number bitch
                                int iCurrentDealnum;
                                string sCurrentDealCell;
                                sCurrentDealCell = DealRows.GetCell(columnStart).ToString();
                                int iStartDeal;
                                int iLastDeal;
                                iStartDeal = sCurrentDealCell.LastIndexOf(" ");
                                iLastDeal = sCurrentDealCell.Length;                               
                                sCurrentDealCell = sCurrentDealCell.Substring(iStartDeal, iLastDeal - iStartDeal).Trim();
                                if(int.TryParse(sCurrentDealCell, out iCurrentDealnum))
                                {
                                    //we got us a mother fucking deal number, lets check the customer quantity for this deal
                                    string sCurrentQty;
                                    sCurrentQty = cRow.GetCell(columnStart).ToString();
                                    //read the cell for our current fucking customer and our current fucking deal
                                    int iCurrentQty;
                                    if (int.TryParse(sCurrentQty, out iCurrentQty))
                                    {
                                        //fuck yeah ass muncher we got a valid quantity
                                        if (iCurrentQty > 0)
                                        {
                                            //the quantity is bigger than 1.. add that shit to the table
                                            DataRow drAdd = dtMaster.NewRow();
                                            drAdd["CustNumber"] = iCustomerNum.ToString();
                                            drAdd["DealNumber"] = iCurrentDealnum.ToString();
                                            drAdd["Quantity"] = iCurrentQty;
                                            dtMaster.Rows.Add(drAdd);
                                            dtMaster.AcceptChanges();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                //return that table for fucks sake!
                return dtMaster;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                return dtMaster;
            }
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            InitializeWorkbook(openFileDialog1.FileName);
            int iCounter = 1;
            
            foreach (ISheet sheet in workBook)
            {
                cboSheets.Items.Add(sheet.SheetName);
                iCounter++;
            }          

            if (iCounter > 1)
            {
                label1.Visible = true;
                cboSheets.Visible = true;
                cboSheets.SelectedIndex = 0;

                lblDescription.Visible = true;
                textBox1.Visible = true;
            }
        }

        

        private void button1_Click(object sender, EventArgs e)
        {
            
            //saveFileDialog1.ShowDialog();
            //clsStringBuilder stringbuilder = new clsStringBuilder();
            //string sCSV = stringbuilder.ToCSV(dataSet1.Tables[0]);

            string description = textBox1.Text;

            string CsvFpath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "FWDEAORD.csv");
            try
            {
                StreamWriter csvFileWriter = new StreamWriter(CsvFpath, false);

                int countColumn = dataGridView1.ColumnCount -1;

                string docust = "DOCUST";
                string doitem = "DOITEM";
                string doqtyo = "DOQTYO";
                string dodesc = "DODESC";

                csvFileWriter.WriteLine(docust + ',' + doitem + ',' + doqtyo + ',' + dodesc);

                foreach (DataGridViewRow dataRowObject in dataGridView1.Rows)
                {
                    if (!dataRowObject.IsNewRow)
                    {
                        string dataFormGrid = "";

                        dataFormGrid = dataRowObject.Cells[0].Value.ToString();

                        for (int i = 1; i <= countColumn; i++)
                        {
                            dataFormGrid = dataFormGrid + ',' + dataRowObject.Cells[i].Value.ToString();
                            
                        }
                        dataFormGrid = dataFormGrid + ',' + description;
                        csvFileWriter.WriteLine(dataFormGrid);
                    }
                }

                csvFileWriter.Flush();
                csvFileWriter.Close();
            }
            catch (Exception exceptionObject)
            {
                MessageBox.Show(exceptionObject.ToString());
            }
        }
    }
}
