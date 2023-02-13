using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.Diagnostics;
using System.Globalization;

namespace MPE_Project
{
    public partial class Form1 : Form
    {
        //-----------------------------------------------------Validate Section---------------------------------------------------------------//
        readonly List<string> files = new();
        readonly Dictionary<string, string> Database = new();
        readonly List<string> Headers = new() { "Supplier Name", "Component Type", "APN", "MPN", "Program Name", "Lot Code", "Date Code", "Test Step", "Tester Platform", "Test Program Name", "Manufacturing Flow", "Date Tested", "Lot Qty", "Yield %", "SYL" };
        bool headCheck, mismatchValues, whiteSpace = false;
        readonly List<string> ErrorsList = new();
        //------------------------------------------------------------------------------------------------------------------------------------//
        //-----------------------------------------------------Generate Section---------------------------------------------------------------//
        readonly Dictionary<string, string> ListAddressBase = new(); //Column name & address (index) for copiable data
        readonly Dictionary<string, string> ListAddressBaseNon = new();
        readonly Dictionary<string, string> ListValueBase = new();
        readonly Dictionary<string, string> ListAddressAuto = new(); //Column name and address for non copiable data
        readonly Dictionary<string, string> LotsAddressAuto = new();
        readonly Dictionary<string, string> BinFRAddressAuto = new();
        bool DateCodeFlag = false;
        //------------------------------------------------------------------------------------------------------------------------------------//
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            files.Clear();
            OpenFileDialog ofd = new()
            {
                Title = "Select Files",
                Multiselect = true,
                Filter = "CSV Files|*.csv*"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                foreach (string path in ofd.FileNames)
                {
                    files.Add(path);
                }
                textBox1.Text = Path.GetFileName(ofd.FileName);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            var progress = new Progress<int>(ProgressBar);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            if (radioButton1.Checked)
            {
                DateCodeFlag = false;
                if(string.IsNullOrEmpty(comboBox1.Text) | string.IsNullOrEmpty(comboBox2.Text) | string.IsNullOrEmpty(comboBox3.Text))
                {
                    label11.Text = "*Load All Files to Start MPE Project*";
                    label11.ForeColor = Color.DarkRed;
                }
                else if (!textBox3.Text.Contains(comboBox1.Text))
                {
                    label11.Text = "*Format File Does Not Match with P/N*";
                    label11.ForeColor = Color.DarkRed;
                }
                else 
                {
                    Debug.WriteLine("Generate MODE");
                    //---------------------------------------------------Load Files---------------------------------------------------------------//
                    //----------------------Create MPE File--------------------------//
                    using var mpePack = new ExcelPackage();
                    var mpeWS = mpePack.Workbook.Worksheets.Add("Sheet1");
                    mpeWS.Cells["A1"].Value = "Hello World!";
                    Debug.WriteLine("MPE File Created!");
                    //----------------------Autocounting File------------------------//
                    using var autoPack = new ExcelPackage(autoPath);
                    var autoWS = autoPack.Workbook.Worksheets[0];
                    Debug.WriteLine("Autocounting File Loaded!");
                    //----------------------.csv Format File-------------------------//
                    using var basePack = new ExcelPackage();
                    var baseWS = basePack.Workbook.Worksheets.Add("Sheet1");
                    if (File.Exists(basePath))
                    {
                        var baseFile = new FileInfo(basePath);
                        var format = new ExcelTextFormat
                        {
                            Delimiter = ','
                        };
                        var ts = TableStyles.Dark1;
                        baseWS.Cells["A1"].LoadFromText(baseFile, format, ts, FirstRowIsHeader: false);
                        Debug.WriteLine("CSV File Loaded!");
                    }
                    //----------------------------------------------------------------------------------------------------------------------------//
                    //-----------------------------------------------------First Row MPE----------------------------------------------------------//
                    baseWS.Cells["A1:CN1"].Copy(mpeWS.Cells["A1"]);
                    //----------------------------------------------------------------------------------------------------------------------------//
                    //---------------------------------------------Get Base Addresses-------------------------------------------------------------//
                    GetBaseAddresses(baseWS);
                    //----------------------------------------------------------------------------------------------------------------------------//
                    //--------------------------------------------------Filtering Autocounting----------------------------------------------------//
                    //Debug.WriteLine("Autocounting Range: " + autoWS.Dimension.Address);

                    Autocounting(autoWS);
                    if (DateCodeFlag)
                    {
                        //----------------------------------------------------------------------------------------------------------------------------//
                        //--------------------------------------------------Filling Out MPE-----------------------------------------------------------//
                        //---------------------------------------Copiable Data-----------------------------------------------//
                        //Base File --> MPE File
                        CopiableData(mpeWS);
                        //--------------------------------Get Autocounting Non Copiable Data---------------------------------//
                        //Auto File --> MPE File
                        NonCopiableData(autoWS, mpeWS);
                        //----------------------------------------------------------------------------------------------------------------------------//
                        //-----------------------------------------------------Close Files------------------------------------------------------------//
                        //----------------------Create MPE File--------------------------//
                        string numb = "";
                        foreach (char number in comboBox2.Text)
                        {
                            if (char.IsNumber(number))
                            {
                                numb = string.Concat(numb, number);
                            }
                        }
                        string path = string.Concat(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "\\", ListValueBase["Program Name"], "_", ListValueBase["MPN"], "_", ListValueBase["APN"], "_WW", numb, "-mpe-raw.csv");
                        label11.Text = "Done!";
                        label11.ForeColor = Color.DarkGreen;
                        End(mpeWS, path);
                        MessageBox.Show("MPE Report Succesfully Done!" + "\n" + "New path: " + path, "Results", MessageBoxButtons.OK);
                        label11.Text = "Start MPE Project!";
                        label11.ForeColor = Color.DarkGreen;
                        //----------------------------------------------------------------------------------------------------------------------------//
                    }
                }
            }
            else
            {
                Debug.WriteLine("Validate MODE");
                // For each file selected
                foreach (string path in files)
                {
                    headCheck = false;
                    mismatchValues = false;
                    whiteSpace = false;
                    ErrorsList.Clear();
                    //-----------------------------------------------Load .CSV Files-------------------------------------------------------//
                    //var MPEws = Start(path);
                    var format = new ExcelTextFormat
                    {
                        Delimiter = ','
                    };
                    var file = new FileInfo(path);
                    using var package = new ExcelPackage();
                    var MPEws = package.Workbook.Worksheets.Add(Path.GetFileName(path)); //var MPEws = package.Workbook.Worksheets.First();
                    var ts = TableStyles.Dark1;
                    MPEws.Cells["A1"].LoadFromText(file, format, ts, FirstRowIsHeader: false);
                    //Debug.WriteLine(MPEws);
                    Debug.WriteLine("CSV File Loaded");
                    MPEws.Cells["A2:CN2"].Copy(MPEws.Cells["A2"]);
                    //Debug.WriteLine(MPEws.Cells["A10"].Value);
                    //---------------------------------------------------------------------------------------------------------------------//
                    //--------------------------------------------------Validate Head Rows-------------------------------------------------//
                    Debug.WriteLine("1: Validate Head Rows");
                    ValidateHeadRow(MPEws);
                    //---------------------------------------------------------------------------------------------------------------------//
                    //----------------------------------------------Get Elements per Column------------------------------------------------//
                    Debug.WriteLine("2. Get Elements per Column");
                    Debug.WriteLine("Columns: " + Convert.ToString(MPEws.Dimension.Columns));
                    Debug.WriteLine("Rows: " + Convert.ToString(MPEws.Dimension.Rows - 1));
                    byte realCol = 0;
                    for (short a = 2; a <= MPEws.Dimension.Rows - 1; a++)
                    {
                        byte fake = 0;
                        foreach (var cell in MPEws.Cells[string.Concat("A", Convert.ToString(a), ":", "CN", Convert.ToString(a))])
                        {
                            if (string.IsNullOrEmpty(Convert.ToString(cell.Value)))
                            {
                                if (realCol <= fake)
                                {
                                    realCol = fake;
                                }
                                break;
                            }
                            fake++;
                        }
                    }
                    Debug.WriteLine("Real Columns: " + realCol);
                    //---------------------------------------------------------------------------------------------------------------------//
                    //----------------------------------------------Matching Values by Columns---------------------------------------------//
                    Debug.WriteLine("3. Matching Same Values");
                    MatchSameValues(MPEws, realCol);
                    //---------------------------------------------------------------------------------------------------------------------//
                    //---------------------------------------------------Check White Spaces------------------------------------------------//
                    Debug.WriteLine("4. Check White Spaces");
                    WhiteSpace(MPEws, realCol);
                    //save file, no need to close in code
                    MPEws.Cells["L:L"].Style.Numberformat.Format = "mm-dd-yyyy";
                    //---------------------------------------------------------------------------------------------------------------------//
                    //---------------------------------------------------Show Results------------------------------------------------------//
                    DisplayChecks(path);
                    //---------------------------------------------------------------------------------------------------------------------//
                    //----------------------------------------------------Save .CSV Files--------------------------------------------------//
                    End(MPEws, path);
                    //---------------------------------------------------------------------------------------------------------------------//
                }
            } 
        }
        private void ValidateHeadRow(ExcelWorksheet MPEws)
        {  
            Database.Clear();
            byte sub_col = 1;
            for (byte col = 1; col <= MPEws.Dimension.Columns; col++)
            {
                string key = MPEws.Cells[1, col].Text;
                string value = MPEws.Cells[2, col].Text;
                if (key.Equals("Supplier Name") | key.Equals("Component Type") | key.Equals("APN") | key.Equals("MPN") | key.Equals("Program Name") | key.Equals("Date Code") | key.Equals("Test Step") | key.Equals("Manufacturing Flow") | key.Equals("SYL"))
                {
                    if (!string.IsNullOrEmpty(value.ToString()))
                    {
                        Database.Add(key.ToString(), value);
                    }
                    else
                    {
                        //white space
                        short count = 3;
                        while (!string.IsNullOrEmpty(value))
                        {
                            value = MPEws.Cells[count, col].Text.ToString();
                            count++;
                        }
                        Database.Add(key.ToString(), value);
                    }
                }
                if (col <= 15 & !Headers.Contains(key))
                {
                    //headers wrong values
                    headCheck = true;
                    Debug.WriteLine("Error at cell: " + 1 + ","+col);
                    ErrorsList.Add(MPEws.Cells[1, col].Address);
                }
                else if (col > 15)
                {
                    if (((key.Contains("_Number") | key.Contains("_Name") | key.Contains("_SBL")) & !string.IsNullOrEmpty(value) & !value.Equals("")))
                    {
                        Database.Add(key, value);
                    }
                    //check BinX_ Columns
                    switch (col%4)
                    {
                        case 0:
                            sub_col++;
                            string BinNumber = string.Concat("BinX", Convert.ToString(sub_col), "_Number");
                            if((col == 92 & !key.Equals("Comment")))
                            {
                                //headers wrong values
                                headCheck = true;
                                //Debug.WriteLine("Error at cell: " + 1 + "," + col);
                                ErrorsList.Add(MPEws.Cells[1, col].Address);
                            }
                            else if (col< 92 & !key.Equals(BinNumber))
                            {
                                //headers wrong values
                                headCheck = true;
                                //Debug.WriteLine("Error at cell: " + 1 + "," + col);
                                ErrorsList.Add(MPEws.Cells[1, col].Address);
                            }
                            break;
                        case 1:
                            string BinName = string.Concat("BinX", Convert.ToString(sub_col), "_Name");
                            if (!key.Equals(BinName))
                            {
                                //headers wrong values
                                headCheck = true;
                                //Debug.WriteLine("Error at cell: " + 1 + "," + col);
                                ErrorsList.Add(MPEws.Cells[1, col].Address);
                            }
                            break;
                        case 2:
                            string Bin_FailRate = string.Concat("BinX", Convert.ToString(sub_col), "_%");
                            if (!key.Equals(Bin_FailRate))
                            {
                                //headers wrong values
                                headCheck = true;
                                //Debug.WriteLine("Error at cell: " + 1 + "," + col);
                                ErrorsList.Add(MPEws.Cells[1, col].Address);
                            }
                            break;
                        case 3:
                            string BinSBL = string.Concat("BinX", Convert.ToString(sub_col), "_SBL");
                            if (!key.Equals(BinSBL))
                            {
                                //headers wrong values
                                headCheck = true;
                                //Debug.WriteLine("Error at cell: " + 1 + "," + col);
                                ErrorsList.Add(MPEws.Cells[1, col].Address);
                            }
                            break;
                    }
                }
            }
            Debug.WriteLine("Process 1 finished!");
            //Console.WriteLine(string.Join(", ", Database.Select(pair => $"{pair.Key} => {pair.Value}")));
        }
        private void MatchSameValues(ExcelWorksheet MPEws, byte realCol)
        {
            for (byte i = 1; i <= realCol; i++)
            {
                if (Database.ContainsKey(MPEws.Cells[1, i].Text) | MPEws.Cells[1, i].Text.Equals(""))
                {
                    for (int j = 2; j <= MPEws.Dimension.Rows - 1; j++)
                    {
                        var cell = MPEws.Cells[j, i].Text;
                        if (!Database.ContainsValue(cell.ToString()))
                        {
                            //mismatch value
                            mismatchValues = true;
                            //Debug.WriteLine("Error at cell: " + MPEws.Cells[j,i].Address);
                            ErrorsList.Add(MPEws.Cells[j, i].Address);

                        }
                    }
                    /*if (mismatchValues)
                    {
                        break;
                    }
                    //PENDING TO IMPLEMENT, SEARCH .All method
                    /*bool check = checkColumn.All(value => value.Equals(Database[FirstRow.ElementAt(i).StringValue]));
                    if (!check)
                    {
                        //One or more elements do not match values for this column
                        System.Diagnostics.Debug.WriteLine("Mega Failed");
                        mistmatchValues = true;
                        break;
                    }*/
                }
            }
            Debug.WriteLine("Process 3 finished!");
        }
        private void WhiteSpace(ExcelWorksheet MPEws, byte realCol)
        {
            for (byte i = 1; i <= realCol; i++)
            {
                for (short j=1; j<=MPEws.Dimension.Rows-1; j++)
                {
                    var cell = MPEws.Cells[j, i].Text;
                    if (string.IsNullOrEmpty(cell))
                    {
                        // white space
                        Debug.WriteLine("White Space at: " + j + i);
                        ErrorsList.Add(MPEws.Cells[1, j].Address);
                        whiteSpace = true;
                        break;

                    }
                }
            }
            Debug.WriteLine("Process 4 finished!");
        }
        private void DisplayChecks(string path)
        {
            string message = "";
            if (!whiteSpace)
            {
                message = string.Concat(message, "Passed White Spaces Check!");
            }
            else
            {
                message = string.Concat(message, "Failed White Spaces Check!");
            }
            if (!mismatchValues)
            {
                message = string.Concat(message, "\n", "Passed Matched Values Check!");
            }
            else
            {
                message = string.Concat(message, "\n", "Failed Matched Values Check!");
            }
            if (!headCheck)
            {
                message = string.Concat(message, "\n", "Passed Header Names Check!");
            }
            else
            {
                message = string.Concat(message, "\n", "Failed Header Names Check!");
            }
            if (ErrorsList.Any())
            {
                message = string.Concat(message, "\n", "Error at Cells: ");
                foreach (string error in ErrorsList)
                {
                    message = string.Concat(message, error, " ");
                }
            }
            label11.Text = "Done!";
            MessageBox.Show(message, string.Concat("Results of ", Path.GetFileName(path)), MessageBoxButtons.OK);
            label11.Text = "Start MPE Project!";

        }
        public static void End(ExcelWorksheet MPEws, string path)
        {
            var formatOut = new ExcelOutputTextFormat
            {
                Delimiter = ',',
            };
            var file = new FileInfo(path);
            MPEws.Cells[1, 1, MPEws.Dimension.Rows, MPEws.Dimension.Columns].SaveToText(file, formatOut);
            Debug.WriteLine("File " + Path.GetFileName(path) + " Closed!");
        }
        //------------------------------------------------------------------------------------------------------------------------------------//
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            label3.Enabled = radioButton2.Checked;
            textBox1.Enabled = radioButton2.Checked;
            button1.Enabled = radioButton2.Checked;

            label4.Enabled = radioButton1.Checked;
            textBox2.Enabled = radioButton1.Checked;
            button3.Enabled = radioButton1.Checked;

            label5.Enabled = radioButton1.Checked;
            textBox3.Enabled = radioButton1.Checked;
            button4.Enabled = radioButton1.Checked;

            label6.Enabled = radioButton1.Checked;
            comboBox1.Enabled = radioButton1.Checked;

            label7.Enabled = radioButton1.Checked;
            comboBox2.Enabled = radioButton1.Checked;

            label10.Enabled= radioButton1.Checked;
            comboBox3.Enabled = radioButton1.Checked;
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            label3.Enabled = radioButton2.Checked;
            textBox1.Enabled = radioButton2.Checked;
            button1.Enabled = radioButton2.Checked;

            label4.Enabled = radioButton1.Checked;
            textBox2.Enabled = radioButton1.Checked;
            button3.Enabled = radioButton1.Checked;

            label5.Enabled = radioButton1.Checked;
            textBox3.Enabled = radioButton1.Checked;
            button4.Enabled = radioButton1.Checked;

            label6.Enabled = radioButton1.Checked;
            comboBox1.Enabled = radioButton1.Checked;

            label7.Enabled = radioButton1.Checked;
            comboBox2.Enabled = radioButton1.Checked;

            label10.Enabled = radioButton1.Checked;
            comboBox3.Enabled = radioButton1.Checked;
        }
        //------------------------------------------------------------------------------------------------------------------------------------//
        //-----------------------------------------------------Generate Section---------------------------------------------------------------//
        string autoPath = "", basePath="";
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new()
            {
                Title = "Select File",
                Filter = "Excel File|*.xlsx*"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                autoPath= ofd.FileName;
                textBox2.Text = Path.GetFileName(ofd.FileName);
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new()
            {
                Title = "Select File",
                Filter = "CSV Files|*.csv*"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                basePath= ofd.FileName;
                textBox3.Text = Path.GetFileName(ofd.FileName);
            }
        }

        private void GetBaseAddresses(ExcelWorksheet baseWS)
        {
            short col = 1;
            ListAddressBase.Clear();
            ListValueBase.Clear();
            ListAddressBaseNon.Clear();
            LotsAddressAuto.Clear();
            BinFRAddressAuto.Clear();
            foreach (var head in baseWS.Cells["A1:CN1"])
            {
                if (head.Text.Equals("Supplier Name") | head.Text.Equals("Component Type") | head.Text.Equals("APN") | head.Text.Equals("Program Name") | head.Text.Equals("Test Step") | head.Text.Equals("Manufacturing Flow") | head.Text.Equals("SYL"))
                {
                    //To copy data
                    ListAddressBase.Add(head.Text, head.Address);
                    ListValueBase.Add(head.Text, baseWS.Cells[2, col].Text);
                }
                if (head.Text.Equals("Date Code"))
                {
                    ListAddressBase.Add(head.Text, head.Address);
                    ListValueBase.Add("Date Code", string.Concat(comboBox2.Text, " ", comboBox3.Text));
                }
                if ((head.Text.Contains("_Number") | head.Text.Contains("_Name") | head.Text.Contains("_SBL")) & !string.IsNullOrWhiteSpace(baseWS.Cells[2, col].Text))
                {
                    ListAddressBase.Add(head.Text, head.Address);
                    ListValueBase.Add(head.Text, baseWS.Cells[2, col].Text);
                }
                if (head.Text.Equals("Lot Code") | head.Text.Equals("Test Program Name") | head.Text.Equals("Tester Platform") | head.Text.Equals("Lot Qty") | head.Text.Equals("Yield %") | (head.Text.Contains("_%") & !string.IsNullOrWhiteSpace(baseWS.Cells[2, col].Text)) | head.Text.Equals("Date Tested"))
                {
                    //Not to copy data 
                    ListAddressBaseNon.Add(head.Text, head.Address);
                }
                if (head.Text.Equals("MPN"))
                {
                    ListAddressBase.Add(head.Text, head.Address);
                    ListValueBase.Add(head.Text, comboBox1.Text);
                }
                col++;
            }
        }
        private void Autocounting(ExcelWorksheet autoWS) 
        {
            short col = 1; short dateCodeAuto = 0, mpnAuto = 0, lotAuto = 0;
            bool flag = false;
            ListAddressAuto.Clear(); LotsAddressAuto.Clear();
            foreach (var head in autoWS.Cells["A1:NH1"])
            {
                if (head.Text.Equals("Lot Code") | head.Text.Equals("Test Program Name") | head.Text.Equals("Tester Platform") | head.Text.Equals("Date Tested") | head.Text.Equals("Lot Qty") | head.Text.Equals("Yield %"))
                {
                    //Not to copy data 
                    ListAddressAuto.Add(head.Text, head.Address);
                }
                if (head.Text.Equals("Date Code"))
                {
                    dateCodeAuto = col;
                }
                else if (head.Text.Equals("MPN"))
                {
                    mpnAuto = col;
                }
                else if (head.Text.Equals("Lot Code"))
                {
                    lotAuto = col;
                }
                col++;
            }
            //-----------------------------------------------Check Date Code---------------------------------------------------------//
            foreach (var cell in autoWS.Cells[1, dateCodeAuto, autoWS.Dimension.Rows, dateCodeAuto])
            {
                if(cell.Text.Equals(string.Concat(comboBox2.Text, " ", comboBox3.Text)))
                {
                    DateCodeFlag = true;
                    break;
                }
            }
            //-----------------------------------------------------------------------------------------------------------------------//
            if (DateCodeFlag)
            {
                for (int row = 2; row <= autoWS.Dimension.Rows; row++)
                {
                    var mpn = autoWS.Cells[row, mpnAuto];
                    var week = autoWS.Cells[row, dateCodeAuto];
                    var lot = autoWS.Cells[row, lotAuto];
                    if (string.Concat(comboBox2.Text, " ", comboBox3.Text).Equals(week.Text) & mpn.Text.Contains(comboBox1.Text))
                    {
                        flag = false;
                        foreach (string key in LotsAddressAuto.Keys)
                        {
                            if (key.Substring(0, key.Length - 2).Equals(lot.Text.Substring(0, lot.Text.Length - 2)))
                            {
                                flag = true;
                                break;
                            }
                        }
                        if (!flag)
                        {
                            LotsAddressAuto.Add(lot.Text, lot.Address);
                        }
                    }
                }
                string binNumber = "";
                BinFRAddressAuto.Clear();
                foreach (var cell in autoWS.Cells["A1:NH1"].Where(a => a.Text.Contains("_%")))
                {
                    binNumber = "";
                    for (int i = 0; i <= cell.Text.Length - 1; i++)
                    {
                        if (char.IsDigit(cell.Text[i]))
                        {
                            binNumber = string.Concat(binNumber, cell.Text[i]);
                        }
                    }
                    //Debug.WriteLine(binNumber);
                    foreach (KeyValuePair<string, string> cell2 in ListValueBase.Where(a => a.Key.Contains("_Number")))
                    {
                        if (binNumber.Equals(cell2.Value))
                        {
                            BinFRAddressAuto.Add(cell.Text, cell.Address);
                            break;
                        }
                    }
                }
            }
            else
            {
                label11.Text = "Date Code does not exist in Autocounting File. Try it again." + "\n Issue with: " + string.Concat(comboBox2.Text, " ", comboBox3.Text);
            }
        }
        private void CopiableData(ExcelWorksheet mpeWS)
        {
            string addressAuto, letterAuto = "";
            foreach (KeyValuePair<string, string> hashmap in ListAddressBase)
            {
                if (ListValueBase.ContainsKey(hashmap.Key))
                {
                    if (hashmap.Value.Length > 2)
                    {
                        letterAuto = hashmap.Value.Substring(0, 2);
                    }
                    else
                    {
                        letterAuto = hashmap.Value.Substring(0, 1);
                    }
                    addressAuto = String.Concat(letterAuto, "2");
                    //mpeWS.Cells[address].Value = ListValueBase[hashmap.Key];
                    //string addresses = String.Concat(letter, "2:", letter, LotsAddressAuto.Count + 1);
                    for (int row = 2; row <= LotsAddressAuto.Count + 1; row++)
                    {
                        addressAuto = String.Concat(letterAuto, row);
                        mpeWS.Cells[addressAuto].Value = ListValueBase[hashmap.Key];
                    }
                }
            }
        }
        private void NonCopiableData(ExcelWorksheet autoWS, ExcelWorksheet mpeWS)
        {
            string letterAuto = "", numberAuto="", addressAuto;
            foreach (KeyValuePair<string, string> hashmap in ListAddressAuto.Concat(BinFRAddressAuto).Where(a => a.Key != "Lot Code"))
            {
                string cutBin = "", letterMPE = "", binNumber = "";
                short rowsBase = 2; 
                for (int i = 0; i <= hashmap.Value.Length - 1; i++)
                {
                    if (char.IsDigit(hashmap.Value[i]))
                    {
                        letterAuto = hashmap.Value.Substring(0, i);
                        break;
                    }
                }
                if (hashmap.Key.Contains("Bin"))
                {
                    // MPE File Address
                    for (int i = 0; i <= hashmap.Key.Length - 1; i++)
                    {
                        if (char.IsDigit(hashmap.Key[i]))
                        {
                            binNumber = string.Concat(binNumber, hashmap.Key[i]);
                        }
                    }
                    foreach (KeyValuePair<string, string> cell in ListValueBase.Where(a => a.Key.Contains("_Number")))
                    {
                        if (cell.Value.Contains(binNumber) & cell.Value.Length <= 3)
                        {
                            for (int i = 0; i <= cell.Key.Length - 1; i++)
                            {
                                if (!char.IsLetterOrDigit(cell.Key[i]))
                                {
                                    cutBin = cell.Key.Substring(0, i);
                                }
                            }
                            foreach (char cell3 in ListAddressBaseNon[string.Concat(cutBin, "_%")])
                            {
                                if (char.IsLetter(cell3))
                                {
                                    letterMPE = string.Concat(letterMPE, cell3);
                                }
                            }
                            break;
                        }
                    }
                }
                else
                {
                    letterMPE = ListAddressBaseNon[hashmap.Key].Substring(0, 1);
                }
                foreach (KeyValuePair<string, string> lot in LotsAddressAuto)
                {
                    for (int i = 0; i <= lot.Value.Length - 1; i++)
                    {
                        if (!char.IsLetter(lot.Value[i]))
                        {
                            numberAuto = lot.Value.Substring(i);
                            break;
                        }
                    }
                    addressAuto = String.Concat(letterAuto, numberAuto);
                    string addressMPE = string.Concat(letterMPE, rowsBase);
                    if (hashmap.Key.Contains("Bin"))
                    {
                        mpeWS.Cells[addressMPE].Value = Math.Round(Convert.ToDouble(autoWS.Cells[addressAuto].Value) * 100, 2);
                    }
                    else
                    {
                        switch (hashmap.Key)
                        {
                            case "Yield %":
                                mpeWS.Cells[addressMPE].Value = Math.Round(Convert.ToDouble(autoWS.Cells[addressAuto].Value), 2);
                                break;
                            case "Date Tested":
                                mpeWS.Cells[addressMPE].Value = autoWS.Cells[addressAuto].Value;
                                mpeWS.Cells[addressMPE].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                break;
                            default:
                                mpeWS.Cells[addressMPE].Value = autoWS.Cells[addressAuto].Value;
                                break;
                        }
                    }
                    rowsBase++;
                }
            }
            letterAuto = ListAddressAuto["Lot Code"].Substring(0, 1);
            string addresses = String.Concat(letterAuto, "2:", letterAuto, LotsAddressAuto.Count + 1);
            mpeWS.Cells[addresses].FillList(LotsAddressAuto.Keys);
        }

        //-----------------------------------------------------Extras Section----------------------------------------------------------------//
        private void ProgressBar(int percentage)
        {
            progressBar1.Value=percentage;
        }
    }
}