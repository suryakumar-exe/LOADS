using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using ExcelDataReader;
using System.Threading;

using System.Reflection;
using Spire.Xls;
using Spire.Xls.Collections;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.IO;
using ExcelDataReader;

using Aspose.Words.Tables;


namespace Excel_Manipulation_Learning
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataTableCollection tableCollection;
        string Fatigue_excel = "";
        string QuickView_excel = "";
        string pitch_excel = "";
        string hub_excel = "";
        string mainshaft_excel = "";
        string gearbox_excel = "";
        string gearboxelickoff_excel = "";
        string gearboxzf_excel = "";
        string blade_excel = "";
        string yaw_excel = "";
        string mainframe_excel = "";
        string tower_excel = "";
        string genframe_excel = "";
        string steeltower_excel = "";
        string towerca_excel = "";
        string additional_sensor_excel = "";

        string refload_excel = "";
        string refload_sub = "";

        int cmt = 0;
        int f = 1;

        int ff = 1;
        int mc = 0;
        int csn = 0;
        int mval = 0;
        int dle = 0;
        int frl = 0;
        int awp = 0;

        int qf = 1;
        int ref_val = 0;
        int qvcsn = 0;
        int qvm = 0;


        int ref_mc = 0;
        int ref_mval = 0;
        int ref_dle = 0;
        int ref_awp = 0;
        int ref_csn = 0;

        int q_ref = 1;

        Excel.Application xlApp1;
        Excel.Workbooks xlWbks1;
        Excel.Workbook xlWbk1;

        Excel.Application xlApp1_rl;
        Excel.Workbooks xlWbks1_rl;
        Excel.Workbook xlWbk1_rl;

        Excel.Range xlUsedRange;

        Excel.Worksheet xlWorkSheet1_rl;



        private void button2_Click(object sender, EventArgs e)
        {

            try
            {
                /*
                                Excel.Application xlApp1 = new Excel.Application();
                                Excel.Workbooks xlWbks1 = xlApp1.Workbooks;
                                Excel.Workbook xlWbk1 = xlWbks1.Open(@"D:\LOADS PROJECT\QuickView_Design_Loads_Delta4000 (004).xlsx");
                                Excel.Worksheet xlWorkSheet1 = xlWbk1.Sheets["Pitch N155_4.X"];
                                xlWorkSheet1.Activate();
                                xlApp1.Visible = true;
                                *//*  string m_del_value = "DEL m=" + v;
                                  string m_meqn_value = "MEQn m=" + v;*//*
                                Excel.Range xlUsedRange1 = xlWorkSheet1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                                //object[] myCriteria = new object[] { "MREQn m=3", "MEQn m=5"};
                                xlUsedRange1.AutoFilter(2, "DEL m=5", Excel.XlAutoFilterOperator.xlAnd, Type.Missing, Type.Missing);
                                //xlUsedRange1.AutoFilter(2, "MREQn m=3", Excel.XlAutoFilterOperator.xlOr,,Type.Missing);

                                Excel.Range filteredRange1 = xlUsedRange1.SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                                MessageBox.Show("Clear All Filter?");
                                xlUsedRange1.Worksheet.AutoFilterMode = false;*/

                xlApp1_rl = new Excel.Application();
                xlWbks1_rl = xlApp1_rl.Workbooks;
                xlWbk1_rl = xlWbks1_rl.Open(refload_excel);
                xlWorkSheet1_rl = xlWbk1_rl.Sheets[refload_sub];
                xlWorkSheet1_rl.Activate();
                xlApp1_rl.Visible = true;
                var refdic = new Dictionary<string, string>();
                Excel.Range xlRange = xlWorkSheet1_rl.UsedRange;
                //Excel.Range xlRange = xlWorksheet.UsedRange;
                xlRange.AutoFilter(7, "10", Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                xlRange.AutoFilter(1, "Blade", Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                int rowCount = xlRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Count;
                int colCount = xlRange.Columns.Count;

                foreach (Excel.Range row in xlRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows)
                {
                    Excel.Range MainComponent = (Excel.Range)row.Cells[1, 24];

                    Excel.Range DIE = (Excel.Range)row.Cells[1, 1];

                    if ((MainComponent.Value2 != null) && (DIE.Value2 != null))
                    {
                        MessageBox.Show(DIE.Value2.ToString());
                        MessageBox.Show(MainComponent.Value2.ToString());
                    }
                    else
                    {
                        break;
                    }
                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background

            }
            catch (Exception theException)
            {
                xlWbk1_rl.Close(false);
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }







        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {


                //Excel File 1:Fatigue Exel
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbooks xlWbks = xlApp.Workbooks;
                Excel.Workbook xlWbk = xlWbks.Open(Fatigue_excel);
                Excel.Sheets xlSheets = xlWbk.Sheets;
                Excel.Worksheet xlWorkSheet = xlSheets.get_Item(1);
                //proggress bar (some error)
                if (!backgroundWorker1.IsBusy)
                {
                    backgroundWorker1.RunWorkerAsync();
                }

                List<KeyValuePair<string, string>> d = new List<KeyValuePair<string, string>>();
                Excel.Range xlUsedRange2 = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                Excel.Range filteredRange4 = xlUsedRange2.SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                int colnum = 1;
                //to find column index of fatigue excel
                foreach (Excel.Range row in filteredRange4.Columns)
                {
                    Excel.Range cell = (Excel.Range)row.Cells[1, 1];
             
                    if (cell.Value2 == "Main Component")
                    {
                        mc = colnum;
                    }
                    else if (cell.Value2 == "Common Sensor Name")
                    {
                        csn = colnum;
                    }
                    else if (cell.Value2 == "m")
                    {
                        mval = colnum;

                    }
                    else if (cell.Value == "D.L.E.")
                    {
                        dle = colnum;

                    }
                    else if (cell.Value2 == "FLAp reference load")
                    {
                        frl = colnum;
                    }
                    else if (cell.Value2 == "FLAp AWP margin")
                    {
                        awp = colnum;
                        break;
                    }
                    colnum += 1;
                }

                //to add a csn and m values (unique entry)
                foreach (Excel.Range area3 in filteredRange4.Areas)
                {
                    foreach (Excel.Range row in area3.Rows)
                    {
                        int flg = 0;
                        if (((Excel.Range)row.Cells[1, mc]).Text == "Main Component" && ((Excel.Range)row.Cells[1, mval]).Text == "m")
                        {

                            continue;
                        }
                        else if ((((Excel.Range)row.Cells[1, mc]).Text == "") || (((Excel.Range)row.Cells[1, mval]).Text == ""))
                        {
                            goto LoopEnd;
                        }


                        else if ((((Excel.Range)row.Cells[1, mc]).Text != "") && (((Excel.Range)row.Cells[1, mval]).Text != ""))
                        {
                            foreach (var q in d)
                            {
                                if ((q.Key == ((Excel.Range)row.Cells[1, mc]).Text) && (q.Value == ((Excel.Range)row.Cells[1, mval]).Text))
                                {
                                    flg += 1;
                                }
                            }
                            if (flg == 0)
                            {

                                d.Add(new KeyValuePair<string, string>(((Excel.Range)row.Cells[1, mc]).Text, ((Excel.Range)row.Cells[1, mval]).Text));
                            }
                        }
                    }
                }
            LoopEnd:
                MessageBox.Show("Finished with calculations.");



                xlUsedRange = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

                int d_count = d.Count;
                //pass the csn and m value one by one
                int x = 0;
                foreach (var q in d)
                {
                    xlUsedRange.AutoFilter(mc, q.Key, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, Type.Missing);
                    xlUsedRange.AutoFilter(mval, q.Value, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, Type.Missing);
                    xlApp.Visible = true;

                    if (x == d_count - 1)
                    {
                        cmt = 1;
                    }

                    x += 1;
                    perform(q.Key, q.Value);



                }


                void perform(string ki, string val)
                {
                    //background task
                    if (!backgroundWorker1.IsBusy)
                    {
                        backgroundWorker1.RunWorkerAsync();
                    }
                    Excel.Range filteredRange = xlUsedRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                    var sensorName1 = new List<string>();
                    var ref_sen = new List<string>();
                    var asn_arr = new List<string>();
                    foreach (Excel.Range area in filteredRange.Areas)
                    {
                        foreach (Excel.Range row in area.Rows)
                        {
                            if (((Excel.Range)row.Cells[1, csn]).Text == "Common Sensor Name")
                            {
                                continue;
                            }
                            else if (((Excel.Range)row.Cells[1, csn]).Text != "")
                            {
                                string sub = ((Excel.Range)row.Cells[1, csn]).Text;
                                int len = sub.Length;
                                int orglen = len - 2;
                                string blade_text = sub.Substring(orglen);
                                if (blade_text.ToLower() == "bt")
                                {
                                    string mainone = sub.Substring(0, orglen) + "74m";
                                    sensorName1.Add(mainone);
                                    ref_sen.Add(sub);
                                }
                                else if (blade_text.ToLower() == "2t")
                                {
                                    string mainone = sub.Substring(0, orglen) + "50m";
                                    sensorName1.Add(mainone);
                                    ref_sen.Add(sub);
                                }
                                else if (blade_text.ToLower() == "hb")
                                {
                                    string mainone = sub.Substring(0, orglen) + "39m";
                                    sensorName1.Add(mainone);
                                    ref_sen.Add(sub);
                                }
                                else if (blade_text.ToLower() == "1t")
                                {
                                    string mainone = sub.Substring(0, orglen) + "24m";
                                    sensorName1.Add(mainone);
                                    ref_sen.Add(sub);
                                }
                                else if (blade_text.ToLower() == "mc")
                                {
                                    string mainone = sub.Substring(0, orglen) + "17m";
                                    sensorName1.Add(mainone);
                                    ref_sen.Add(sub);
                                }
                                else if (blade_text.ToLower() == "br")
                                {
                                    string mainone = sub.Substring(0, orglen) + "0m";
                                    sensorName1.Add(mainone);
                                    ref_sen.Add(sub);
                                }
                                else
                                {
                                    sensorName1.Add(((Excel.Range)row.Cells[1, csn]).Text);
                                    ref_sen.Add(sub);
                                }


                            }
                            else
                            {
                                break;
                            }

                        }
                    }
                    //open a fatigue file for once
                    if (f == 1)
                    {
                        xlApp1 = new Excel.Application();
                        xlWbks1 = xlApp1.Workbooks;
                        xlWbk1 = xlWbks1.Open(QuickView_excel);
                        f = 0;
                    }

                    object misValue = System.Reflection.Missing.Value;
                    string temp_sheet = "";
                    //to find a appropriate excel file
                    if (ki.Substring(0, 3).ToLower() == "pit")
                    {
                        temp_sheet = pitch_excel;
                    }

                    else if (ki.Substring(0, 3).ToLower() == "hub")
                    {
                        temp_sheet = hub_excel;
                    }
                    else if (ki.ToLower() == "main shaft")
                    {

                        temp_sheet = mainshaft_excel;
                    }
                    else if (ki.Substring(0, 3).ToLower() == "mai" || ki.Substring(0, 3).ToLower() == "rot")
                    {
                        temp_sheet = mainframe_excel;
                    }
                    else if (ki.ToLower() == "gearbox - winergy")
                    {
                        temp_sheet = gearbox_excel;
                    }
                    else if (ki.ToLower() == "gearbox - eickhoff")
                    {
                        temp_sheet = gearboxelickoff_excel;
                    }
                    else if (ki.ToLower() == "gearbox - zf")
                    {
                        temp_sheet = gearboxzf_excel;
                    }
                    else if (ki.Substring(0, 3).ToLower() == "bla")
                    {
                        temp_sheet = blade_excel;
                    }
                    else if (ki.Substring(0, 3).ToLower() == "yaw")
                    {
                        temp_sheet = yaw_excel;
                    }
                    else if (ki.Substring(0, 3).ToLower() == "ste")
                    {
                        temp_sheet = steeltower_excel;
                    }
                    else if (ki.Substring(0, 3).ToLower() == "tow" || ki.ToLower() == "concrete tower")
                    {
                        temp_sheet = tower_excel;
                    }
                    else if (ki.ToLower() == "concrete tower cast adapter")
                    {
                        temp_sheet = towerca_excel;
                    }
                    else if (ki.Substring(0, 3).ToLower() == "gen")
                    {
                        temp_sheet = genframe_excel;
                    }

                    Excel.Worksheet xlWorkSheet1 = xlWbk1.Sheets[temp_sheet];
                    xlWorkSheet1.Activate();
                    int qvcolnum = 1;

                    //to find a column index of reference file
                    if (qf == 1)
                    {
                        foreach (Excel.Range row in xlWorkSheet1.Columns)
                        {
                            Excel.Range cell = (Excel.Range)row.Cells[2, 1];
                            if (cell.Value2 == "Sensor")
                            {
                                qvcsn = qvcolnum;
                                qvm = qvcolnum + 1;
                            }
                            else if (cell.Value2 == "Reference")
                            {
                                ref_val = qvcolnum;
                                qf = 0;
                                break;
                            }

                            qvcolnum += 1;
                        }
                    }
                    string m_del_value = "DEL m=" + val;
                    string m_meqn_value = "MEQn m=" + val;
                    //to filter a value of qv file
                    Excel.Range xlUsedRange1 = xlWorkSheet1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    if (val.Equals("3") && ki.Substring(0, 3).ToLower() != "yaw")
                    {
                        xlUsedRange1.AutoFilter(2, "MREQn m=3", Excel.XlAutoFilterOperator.xlAnd, Type.Missing, Type.Missing);
                    }

                    else if (val.Equals("3.3333"))
                    {
                        xlUsedRange1.AutoFilter(2, "MREQn m=3.3", Excel.XlAutoFilterOperator.xlOr, "MREQn m=3.333", true);
                    }
                    else
                    {
                        xlUsedRange1.AutoFilter(2, m_del_value, Excel.XlAutoFilterOperator.xlOr, m_meqn_value, true);
                    }

                    Excel.Range filteredRange1 = xlUsedRange1.SpecialCells(Excel.XlCellType.xlCellTypeVisible);


                    //to open a reference excel for once
                    if (ff == 1)
                    {
                        xlApp1_rl = new Excel.Application();
                        xlWbks1_rl = xlApp1_rl.Workbooks;
                        xlWbk1_rl = xlWbks1_rl.Open(refload_excel);
                        ff = 0;
                    }

                    xlWorkSheet1_rl = xlWbk1_rl.Sheets[refload_sub];
                    xlWorkSheet1_rl.Activate();
                    xlApp1_rl.Visible = true;


                    Excel.Range xlRange = xlWorkSheet1_rl.UsedRange;
                    xlRange.Columns.Hidden = false;
                    xlRange.AutoFilter(7, val, Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                    xlRange.AutoFilter(1, ki, Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);



                    //finding rowindex
                    int refnum = 1;
                    if (q_ref == 1)
                    {
                        foreach (Excel.Range row in xlRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Columns)
                        {
                            Excel.Range cell = (Excel.Range)row.Cells[1, 1];
                            if (cell.Value2 == "Main Component")
                            {
                                ref_mc = refnum;
                            }
                            else if (cell.Value2 == "m")
                            {
                                ref_mval = refnum;

                            }
                            else if (cell.Value == "D.L.E.")
                            {
                                ref_dle = refnum;

                            }
                            else if (cell.Value2 == "AWP Safety Margin")
                            {
                                ref_awp = refnum;
                            }
                            else if (cell.Value2 == "Common Name")
                            {
                                ref_csn = refnum;
                                q_ref = 0;
                                break;
                            }
                            refnum += 1;
                        }
                    }

                    int rowCount = xlRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Count;
                    int colCount = xlRange.Columns.Count;

                    var refdic = new Dictionary<string, string>();
                    var refms = new Dictionary<string, string>();
                    //to find a del and awp from refernce lload file
                    foreach (Excel.Range row in xlRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows)
                    {
                        Excel.Range MainComponent = (Excel.Range)row.Cells[1, ref_mc];

                        Excel.Range DIE = (Excel.Range)row.Cells[1, ref_dle];
                        Excel.Range SCN = (Excel.Range)row.Cells[1, ref_csn];
                        Excel.Range AWP = (Excel.Range)row.Cells[1, ref_awp];
                        if (((Excel.Range)row.Cells[1, ref_mc]).Text == "Main Component" && ((Excel.Range)row.Cells[1, ref_csn]).Text == "Common Name" && ((Excel.Range)row.Cells[1, ref_dle]).Text == "D.L.E.")
                        {
                            continue;
                        }
                        else if ((MainComponent.Value2 == ki))
                        {
                            refdic.Add(((Excel.Range)row.Cells[1, ref_csn]).Text, ((Excel.Range)row.Cells[1, ref_dle]).Text);
                            refms.Add(((Excel.Range)row.Cells[1, ref_csn]).Text, ((Excel.Range)row.Cells[1, ref_awp]).Text);
                        }
                        else
                        {
                            goto RefLoopEnd;
                        }
                    }
                RefLoopEnd:
                    label16.Text = ki + " completed";

                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                    //final reference values
                    var reforgsen = new List<string>();
                    for (int i = 0; i < ref_sen.Count; i++)
                    {
                        int flag = 0;
                        foreach (var k in refdic)
                        {
                            if (ref_sen[i].Equals(k.Key, StringComparison.OrdinalIgnoreCase))
                            {


                                reforgsen.Add(k.Value);
                                flag = 1;
                            }
                        }
                        if (flag == 0)
                        {
                            reforgsen.Add("");
                        }

                    }

                    //final awp values
                    var refms_2 = new List<string>();
                    for (int i = 0; i < ref_sen.Count; i++)
                    {
                        int flag = 0;
                        foreach (var k in refms)
                        {

                            if (ref_sen[i].Equals(k.Key, StringComparison.OrdinalIgnoreCase))
                            {
                                refms_2.Add(k.Value);
                                flag = 1;
                            }
                        }
                        if (flag == 0)
                        {
                            refms_2.Add("");
                        }

                    }

                    //to add a csn and dle in dictionary
                    var Dicval = new Dictionary<string, string>();

                    foreach (Excel.Range area1 in filteredRange1.Areas)
                    {
                        foreach (Excel.Range row in area1.Rows)
                        {
                            int SensorRowIndex = 0;


                            if (((Excel.Range)row.Cells[1, qvcsn]).Text != "")
                            {
                                SensorRowIndex = 1;
                            }
                            else
                            {
                                SensorRowIndex = 0;
                            }


                            if (((Excel.Range)row.Cells[1, 1]).Text == "Design Loads Delta4000 > Pitch N155/4.X"
                                || ((Excel.Range)row.Cells[1, 1]).Text == "Design Loads Delta4000 > Hub N149/5.X"
                                || ((Excel.Range)row.Cells[1, 1]).Text == "Design Loads Delta4000 > Main(shaft/bearing) (N133/N149/N155)/4.X (N149/N163)/5.X"
                                || ((Excel.Range)row.Cells[1, 1]).Text == "Design Loads Delta4000 > GBx (N133/N149/N155)/4.X Winergy PZAB3600"
                                || ((Excel.Range)row.Cells[1, 1]).Text == "Design Loads Delta4000 > NR77.5-1 N155/4.X")
                            {
                                continue;
                            }
                            else if (((Excel.Range)row.Cells[SensorRowIndex, qvcsn]).Text != "" && ((Excel.Range)row.Cells[1, qvm]).Text == m_del_value)
                            {
                                Dicval.Add(((Excel.Range)row.Cells[SensorRowIndex, qvcsn]).Text, ((Excel.Range)row.Cells[1, ref_val]).Text);



                            }
                            else if (((Excel.Range)row.Cells[SensorRowIndex, qvcsn]).Text != "" && ((Excel.Range)row.Cells[1, qvm]).Text == m_meqn_value)
                            {

                                String s = ((Excel.Range)row.Cells[SensorRowIndex, qvcsn]).Text + "_MEqn";
                                Dicval.Add(s, ((Excel.Range)row.Cells[1, ref_val]).Text);


                            }
                            else if (((Excel.Range)row.Cells[SensorRowIndex, qvcsn]).Text != "" && ((Excel.Range)row.Cells[1, qvm]).Text == "MREQn m=3")
                            {
                                String s = ((Excel.Range)row.Cells[1, qvcsn]).Text + "_MREqn";
                                Dicval.Add(s, ((Excel.Range)row.Cells[1, ref_val]).Text);
                            }
                            else if (((Excel.Range)row.Cells[SensorRowIndex, qvcsn]).Text != "" && ((Excel.Range)row.Cells[1, qvm]).Text == "MREQn m=3.3")
                            {
                                String s = ((Excel.Range)row.Cells[SensorRowIndex, qvcsn]).Text + "_MREqn";
                                Dicval.Add(s, ((Excel.Range)row.Cells[1, ref_val]).Text);
                            }
                            else
                            {
                                break;
                            }

                        }
                    }

                    //final dle values

                    var orgSen = new List<string>();
                    for (int i = 0; i < sensorName1.Count; i++)
                    {
                        int flag = 0;
                        foreach (var k in Dicval)
                        {
                            if (sensorName1[i].Equals(k.Key, StringComparison.OrdinalIgnoreCase))
                            {
                                orgSen.Add(k.Value);
                                flag = 1;
                            }
                        }
                        if (flag == 0)
                        {
                            orgSen.Add("");
                        }

                    }

                    //fill all the awp,dle and refernece values in fatigue excel
                    int cnt = orgSen.Count;
                    int j = 0;
                    foreach (Excel.Range area in filteredRange.Areas)
                    {
                        foreach (Excel.Range row in area.Rows)
                        {
                            if (((Excel.Range)row.Cells[1, dle]).Text == "D.L.E." || ((Excel.Range)row.Cells[1, frl]).Text == "FLAp reference load" || ((Excel.Range)row.Cells[1, awp]).Text == "FLAp AWP margin")
                            {
                                continue;
                            }
                            else if (((Excel.Range)row.Cells[1, dle]).Text == "" && ((Excel.Range)row.Cells[1, frl]).Text == "" && ((Excel.Range)row.Cells[1, awp]).Text == "")
                            {
                                if (j != (cnt))
                                {
                                    row.Cells[1, dle].value2 = orgSen[j];
                                    row.Cells[1, frl].value2 = reforgsen[j];
                                    row.Cells[1, awp].value2 = refms_2[j];//no not 2 its [1,11]
                                    j += 1;
                                }
                                else
                                {
                                    break;
                                }
              
                            }
                            else
                            {
                                break;
                            }

                        }
                    }
               



                    //to clear all the stored values for next iteration
                    sensorName1.Clear();
                    ref_sen.Clear();
                    orgSen.Clear();
                    Dicval.Clear();
                    reforgsen.Clear();
                    refdic.Clear();
                    refms.Clear();
                    refms_2.Clear();

                    //for process complete indication
                    if (cmt == 1)
                    {
                        MessageBox.Show("process completed");
                        xlUsedRange.Worksheet.AutoFilterMode = false;
                        xlWbk1.Close(false);
                    }

                }


            }
            catch (Exception theException)
            {
                xlWbk1.Close(false);
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }


        }


        //for fatique excel
        private void button3_Click_1(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbookj|*.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Fatigue.Text = openFileDialog.FileName;
                    Fatigue_excel = openFileDialog.FileName;
                }
            }
        }
        //for quickview excel
        private void button4_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbookj|*.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    QuickView.Text = openFileDialog.FileName;
                    //MessageBox.Show(openFileDialog.FileName);
                    QuickView_excel = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.ReadWrite))
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            quickview_cb.Items.Clear();
                            foreach (System.Data.DataTable table in tableCollection)
                            {
                                quickview_cb.Items.Add(table.TableName);
                                hub_cb.Items.Add(table.TableName);
                                mainshaft_cb.Items.Add(table.TableName);
                                gearbox_cb.Items.Add(table.TableName);
                                gbelickoff_cb.Items.Add(table.TableName);
                                gbzf_cb.Items.Add(table.TableName);
                                blade_cb.Items.Add(table.TableName);
                                yaw_cb.Items.Add(table.TableName);
                                mainframe_cb.Items.Add(table.TableName);
                                tower_cb.Items.Add(table.TableName);
                                steeltower_cb.Items.Add(table.TableName);
                                towerca_cb.Items.Add(table.TableName);
                                genframe_cb.Items.Add(table.TableName);
                                gpsupport.Items.Add(table.TableName);
                                roterbearing.Items.Add(table.TableName);
                            }

                        }
                    }
                }
            }

        }

        private void quickview_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable new_one_sheet = tableCollection[quickview_cb.SelectedItem.ToString()];
            pitch_excel = quickview_cb.SelectedItem.ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void hub_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable new_one_sheet1 = tableCollection[hub_cb.SelectedItem.ToString()];
            hub_excel = hub_cb.SelectedItem.ToString();
        }

        private void mainshaft_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable new_one_sheet3 = tableCollection[mainshaft_cb.SelectedItem.ToString()];

            mainshaft_excel = mainshaft_cb.SelectedItem.ToString();
        }

        private void gearbox_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable new_one_sheet1 = tableCollection[gearbox_cb.SelectedItem.ToString()];
            gearbox_excel = gearbox_cb.SelectedItem.ToString();
        }

        private void gbelickoff_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable new_one_sheet1 = tableCollection[gbelickoff_cb.SelectedItem.ToString()];
            gearboxelickoff_excel = gbelickoff_cb.SelectedItem.ToString();
        }

        private void gbzf_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable new_one_sheet1 = tableCollection[gbzf_cb.SelectedItem.ToString()];
            gearboxzf_excel = gbzf_cb.SelectedItem.ToString();
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void blade_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable new_one_sheet1 = tableCollection[blade_cb.SelectedItem.ToString()];
            blade_excel = blade_cb.SelectedItem.ToString();
        }

        private void yaw_cb_SelectedIndexChanged(object sender, EventArgs e)
        {

            System.Data.DataTable new_one_sheet1 = tableCollection[yaw_cb.SelectedItem.ToString()];
            yaw_excel = yaw_cb.SelectedItem.ToString();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable new_one_sheet1 = tableCollection[mainframe_cb.SelectedItem.ToString()];
            mainframe_excel = mainframe_cb.SelectedItem.ToString();
        }

        private void tower_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable new_one_sheet1 = tableCollection[tower_cb.SelectedItem.ToString()];
            tower_excel = tower_cb.SelectedItem.ToString();

        }

        private void genframe_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable new_one_sheet1 = tableCollection[genframe_cb.SelectedItem.ToString()];
            genframe_excel = genframe_cb.SelectedItem.ToString();

        }

        private void steeltower_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable new_one_sheet1 = tableCollection[steeltower_cb.SelectedItem.ToString()];
            steeltower_excel = steeltower_cb.SelectedItem.ToString();

        }

        private void towerca_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable new_one_sheet1 = tableCollection[towerca_cb.SelectedItem.ToString()];
            towerca_excel = towerca_cb.SelectedItem.ToString();
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void gearbox_cb_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            int sum = 0;
            for (int i = 0; i <= 100; i++)
            {
                Thread.Sleep(1000);
                sum = sum + 1;
                backgroundWorker1.ReportProgress(i);
                if (backgroundWorker1.CancellationPending)
                {
                    e.Cancel = true;
                    backgroundWorker1.ReportProgress(0);
                    return;
                }
            }
            e.Result = sum;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            label14.Text = e.ProgressPercentage.ToString() + "%";
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                label14.Text = "Operation fineshed";
            }

        }
        //for reference excel
        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbookj|*.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    refload_main.Text = openFileDialog.FileName;
                    //MessageBox.Show(openFileDialog.FileName);
                    refload_excel = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.ReadWrite))
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            quickview_cb.Items.Clear();
                            foreach (System.Data.DataTable table in tableCollection)
                            {
                                refload_cb.Items.Add(table.TableName);

                            }

                        }
                    }
                }
            }
        }

        private void refload_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable new_one_sheet1 = tableCollection[refload_cb.SelectedItem.ToString()];
            //dataGridView1.DataSource = new_one_sheet.DefaultView;
            refload_sub = refload_cb.SelectedItem.ToString();
        }

        private void refload_main_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void roterbearing_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable new_one_sheet1 = tableCollection[roterbearing.SelectedItem.ToString()];
            //dataGridView1.DataSource = new_one_sheet.DefaultView;
            mainframe_excel = roterbearing.SelectedItem.ToString();
        }
        //for additional sensor excel
        private void btn_asn_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbookj|*.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    cbx_asn.Text = openFileDialog.FileName;
                    //MessageBox.Show(openFileDialog.FileName);
                    additional_sensor_excel = openFileDialog.FileName;
                }
            }
        }

        private void Fatigue_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
