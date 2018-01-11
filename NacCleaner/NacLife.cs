using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace NacCleaner {
    internal class NacLife {

        string fileName = "";
        public List<string> pages;
        public List<Entry> entries;
        public Microsoft.Office.Interop.Excel.Application oXL;
        public Microsoft.Office.Interop.Excel._Workbook oWB;
        public Microsoft.Office.Interop.Excel._Worksheet oSheet;
        public Microsoft.Office.Interop.Excel.Range oRng;
        public object misvalue = System.Reflection.Missing.Value;
        public List<string> pdfLines;
        public int pageNum = 0;
        public int lineNum = 0;
        
        public NacLife(string inFile) {
            pdfLines = new List<string>();
            pages = new List<string>();
            entries = new List<Entry>();
            fileName = System.IO.Path.GetFileName(inFile);

            try {
                //StringBuilder text = new StringBuilder();
                PdfReader pdfReader = new PdfReader(inFile);
                for (int page = 1; page <= pdfReader.NumberOfPages; page++) {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
                    //text.Append(System.Environment.NewLine);
                    //text.Append("\n Page Number:" + page);
                    //text.Append(System.Environment.NewLine);
                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8,
                        Encoding.Default.GetBytes(currentText)));
                    //text.Append(currentText);
                    //pdfReader.Close();
                    pages.Add(currentText);
                }
                int i = 0;

                while (i < pages.Count) {
                    lineNum = 0;
                    Console.WriteLine("\nPage " + (++pageNum));

                    if (pages[i].Length == 0) {
                        pages.RemoveAt(i);
                        Console.WriteLine("Found Empty Page");
                    }
                    else if (pages[i].Contains("A G E N T S U M M A R Y P A G E")) {
                        pages.RemoveAt(i);
                        Console.WriteLine("Found Summary Page");
                    }
                    else if (pages[i].Contains("A N N U A L I Z E D")) {
                        string[] aLines = pages[i].Split('\n');
                        foreach (string aLine in aLines) {
                            Console.Write("Line " + (++lineNum) + ", ");
                            ALineLife tempAL = getALineLife(aLine.Trim());

                            if (tempAL != null) {
                                //System.out.println(tempAL);
                                bool found = false;
                                for (int j = entries.Count - 1; j >= 0; j--) {
                                    if (entries[j].getName() == tempAL.name && entries[j].getPolicyNum() == tempAL.policyNum) {
                                        entries[j].addALineLife(tempAL);
                                        found = true;
                                        break;
                                    }
                                }

                                if (!found) {
                                    Console.WriteLine("WARNING: COULD NOT FIND MATCH, DUMPING AL and matching on last policy match\n\n" + tempAL);
                                    for (int j = entries.Count - 1; j >= 0; j--) {
                                        if (entries[j].getPolicyNum() == tempAL.policyNum) {
                                            entries[j].addALineLife(tempAL);
                                            found = true;
                                            break;
                                        }
                                    }
                                }
                                if (!found)
                                    Console.WriteLine("ERROR: Cannot find any match by policy for annualized line.\nManual intervention requied");
                            }
                        }
                        i++;
                    }
                    else if (pages[i].Contains("C O M M I S S I O N S T A T E M E N T")) {
                        string[] cLines = pages[i].Split('\n');
                        foreach (string cLine in cLines) {
                            Console.Write("Line " + (++lineNum) + ", ");
                            CLineLife tempCL = getCLineLife(cLine.Trim());
                            if (tempCL != null) {
                                bool found = false;
                                foreach (Entry e in entries) {
                                    if (e.getName() == tempCL.name && e.getPolicyNum() == tempCL.policyNum && e.getRate() == tempCL.cRate * 100) {
                                        e.addCLineLife(tempCL);
                                        found = true;
                                        break;
                                    }
                                }

                                if (!found) {
                                    Entry tempE = new Entry(tempCL.name, tempCL.policyNum, tempCL.cRate);
                                    tempE.addCLineLife(tempCL);
                                    entries.Add(tempE);
                                }
                            }
                        }
                        i++;
                    }
                    else {
                        Console.WriteLine("WARNING: page not found: Dumping:\n\n");
                        Console.WriteLine(pages[i]);
                        i++;
                    }//end if
                }//end while

                Console.WriteLine("Finished Scanning for empty pages");
                List<Entry> annuals = new List<Entry>();
                foreach (Entry e in entries) {
                    e.printOut();
                    if (e.getALineLifes().Count > 0) {
                        annuals.Add(e);
                    }
                }
                entries.RemoveAll(entry => entry.getCommissionTotal() == 0);//remove zeros
                CheckIssueDates();
                writeToExcel();
            }
            catch (Exception e) {
                Console.WriteLine(e);
            }
        }

        public ALineLife getALineLife(String s) {
            ALineLife al = null;
            String[] tokens = null;
            String name = "";
            String policyNum = null;
            String accDate = "";
            String issueDate = "";
            int mopd = 0;
            double beginBal = 0.0;
            double currAdv = 0.0;
            double commApp = 0.0;
            double chargeBack = 0.0;
            double endBal = 0.0;
            bool stop = false;
            int pIndex = -1;

            tokens = s.Split(new Char[0], StringSplitOptions.RemoveEmptyEntries);
            if (stop)
                Console.WriteLine("should be bere");
            for (int i = 0; i < tokens.Length; i++) {
                if (tokens[i].StartsWith("LB") || tokens[i].StartsWith("L0")) {
                    policyNum = tokens[i];
                    stop = true;
                    pIndex = i;
                    break;
                }
            }

            if (policyNum == null || policyNum == "") {
                //System.err.println("Couldn't find policy number on comm line");
                return null;
            }

            bool twoNames = false;
            bool threeName = false;
            name = "*" + tokens[0];
            if (!tokens[1].StartsWith("LB") && !tokens[1].StartsWith("L0")) {
                twoNames = true;
                name = name += (" " + tokens[1]);
            }

            if (twoNames && !tokens[2].StartsWith("LB") && !tokens[2].StartsWith("L0")) {
                threeName = true;
                name = name += (" " + tokens[2]);
            }
            Regex nameReg = new Regex(@"[^a-zA-Z\s]");
            name = nameReg.Replace(name, "").Trim();

            if (name.Length < 3)
                Console.WriteLine("Bad Name");

            bool foundAcct = false;

            String regexStr = "^[0-3]?[0-9]/[0-3]?[0-9]/(?:[0-9]{2})?[0-9]{2}$";
            Regex regex = new Regex(regexStr);
            int index = -1;
            int lastDate = -1;

            for (int i = 0; i < tokens.Length; i++) {
                if (regex.Match(tokens[i]).Success) {
                    if (foundAcct == false) {
                        foundAcct = true;
                        accDate = tokens[i];
                        index = i;
                    }
                    else {
                        issueDate = tokens[i];
                        index = i;
                        break;
                    }
                }
            }
            lastDate = index;
            index++;
            mopd = Convert.ToInt32(tokens[index]);
            index++;
            beginBal = Convert.ToDouble(tokens[index].Replace("$", "").Replace("-", ""));
            index++;
            currAdv = Convert.ToDouble(tokens[index].Replace("$", "").Replace("-", ""));
            index++;
            commApp = Convert.ToDouble(tokens[index].Replace("$", "").Replace("-", ""));
            index++;
            chargeBack = Convert.ToDouble(tokens[index].Replace("$", "").Replace("-", ""));
            index++;
            endBal = Convert.ToDouble(tokens[index].Replace("$", "").Replace("-", ""));

            al = new ALineLife(name, policyNum, accDate, issueDate, mopd, beginBal, currAdv, commApp, chargeBack, endBal);
            return al;
        }


        public void writeToExcel() {
            string outFile = "";
            try {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = false;
                oXL.UserControl = false;
                oXL.DisplayAlerts = false;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "Policy";
                oSheet.Cells[1, 2] = "Fullname";
                oSheet.Cells[1, 3] = "Plan";
                oSheet.Cells[1, 4] = "Issue Date";
                oSheet.Cells[1, 5] = "Premium";
                oSheet.Cells[1, 6] = "Rate %";
                oSheet.Cells[1, 7] = "Rate";
                oSheet.Cells[1, 8] = "Commission";
                oSheet.Cells[1, 9] = "Renewal";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "I1").Font.Bold = true;
                oSheet.get_Range("A1", "I1").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                for (int i = 0; i < entries.Count; i++) {
                    object[] outPut = entries[i].getOutput();
                    oSheet.get_Range("A" + (i + 2), "I" + (i + 2)).Value2 = outPut;
                }
                oRng = oSheet.get_Range("A1", "I1");
                oRng.EntireColumn.AutoFit();
                oXL.Visible = false;
                oXL.UserControl = false;

                outFile = GetSavePath();

                oWB.SaveAs(outFile,
                    56, //Seems to work better than default excel 16
                    Type.Missing,
                    Type.Missing,
                    false,
                    false,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing);

                //System.Diagnostics.Process.Start(outFile);
            }
            catch (Exception ex) {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message, "Error");
            }
            finally {
                if (oWB != null)
                    oWB.Close();
                if (File.Exists(outFile))
                    System.Diagnostics.Process.Start(outFile);
            }
        }

        public string GetSavePath() {

            System.Windows.Forms.SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog1.InitialDirectory = "H:\\Desktop\\";
            saveFileDialog1.Filter = "xls|*.xls";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = fileName.Replace(".pdf", "_out");
            
            if (saveFileDialog1.ShowDialog() == DialogResult.OK) {
                return saveFileDialog1.FileName;
            }
            //else System.Windows.Application.Exit();
            return "";
        }

        public int CheckIssueDates() {
            int cnt = 0;
            SqlConnection cs = new SqlConnection("Data Source=RALIMSQL1\\RALIM1; " +
                "Initial Catalog = CAMSRALFG; " +
                "Integrated Security = SSPI; " +
                "Persist Security Info = false; " +
                "Trusted_Connection = Yes");
            SqlCommand cmd = new SqlCommand();
            SqlDataReader reader;
            string currPol = "";

            foreach (Entry entry in entries) {
                if (entry.getIssueDate() == null || entry.getIssueDate() == "") {
                    currPol = entry.getPolicyNum().ToString();
                    string query = @"SELECT Convert(varchar(10),MIN(Sales.IssueDate),101) FROM Sales WHERE Sales.[Policy#]='" + currPol + "';";

                    try {
                        cmd.CommandText = query;
                        cmd.CommandType = System.Data.CommandType.Text;
                        cmd.Connection = cs;
                        cs.Open();

                        reader = cmd.ExecuteReader();

                        if (reader.HasRows) {
                            if (!reader.Read()) {
                                throw new System.Exception("Problem reading results.");
                            }
                            entry.setIssueDate(reader.GetString(0));
                        }
                        else {
                            throw new System.Exception("Couldn't read data from Database or results were empty.");
                        }
                        cnt++;
                    }
                    catch (Exception eIDate) {
                        MessageBox.Show("Couldn't fetch missing issue date for " + currPol + "\n" + eIDate.ToString());
                    }
                    finally {
                        cs.Close();
                    }
                }
            }
            return cnt;
        }

        public CLineLife getCLineLife(string s) {
            CLineLife cl = null;

            string[] lines = s.Split('\n');
            string[] tokens = null;
            string name = null;
            string policyNum = null;
            string type = null;
            string plan = "";
            string accDate = null;
            string dueDate = null;
            int mopd = 0;
            double premium = 0.0;
            double cRate = 0.0;
            double split = 0.0;
            double comm = 0.0;
            bool stop = false;
            int pIndex = -1;

            tokens = s.Split(new Char[0], StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < tokens.Length; i++) {
                if (tokens[i].StartsWith("LB") || tokens[i].StartsWith("L0")) {
                    policyNum = tokens[i];
                    stop = true;
                    pIndex = i;
                    break;
                }
            }

            if (policyNum == null || policyNum == "") {
                //System.err.println("Couldn't find policy number on comm line");
                return null;
            }

            bool twoNames = false;
            bool threeName = false;

            name = tokens[0];

            if (!tokens[1].StartsWith("LB") && !tokens[1].StartsWith("L0")) {
                twoNames = true;
                name = name += (" " + tokens[1]);
            }

            if (twoNames && !tokens[2].StartsWith("LB") && !tokens[2].StartsWith("L0")) {
                threeName = true;
                name = name += (" " + tokens[2]);
            }
            Regex nameReg = new Regex(@"[^a-zA-Z\s]");
            name = nameReg.Replace(name, "").Trim();

            if (name.Length < 3)
                Console.WriteLine("Bad Name");

            bool foundAcct = false;
            bool foundDue = false;

            string regexStr = "^[0-3]?[0-9]/[0-3]?[0-9]/(?:[0-9]{2})?[0-9]{2}$";
            Regex regex = new Regex(regexStr);

            int index = -1;
            int lastDate = -1;

            for (int i = 0; i < tokens.Length; i++) {
                if (regex.Match(tokens[i]).Success) {
                    if (foundAcct == false) {
                        foundAcct = true;
                        accDate = tokens[i];
                        index = i;
                    }
                    else {
                        foundDue = true;
                        dueDate = tokens[i];
                        index = i;
                        break;
                    }
                }
            }
            lastDate = index;
            index++;
            mopd = Convert.ToInt32(tokens[index]);
            index++;
            premium = Convert.ToDouble(tokens[index].Replace("$", "").Replace("-", ""));
            if (tokens[index].Contains("-")) {
                premium = -1 * Math.Abs(premium);
            }
            index++;
            cRate = Convert.ToDouble(tokens[index]);
            index++;

            if (tokens.Length > 10) {
                if (tokens.Length - 1 != index && Convert.ToDouble(tokens[index]) < 1) {
                    split = Convert.ToDouble(tokens[index]);
                    index++;
                }
                comm = Convert.ToDouble(tokens[index].Replace("$", "").Replace("-", ""));
                if (tokens[index].Contains("-")) {
                    comm = -1 * Math.Abs(comm);
                }
            }
            type = tokens[pIndex + 1];
            for (int i = pIndex + 2; i < lastDate - 1; i++) {
                if (tokens[i] != null)
                    plan = plan + " " + tokens[i];
            }
            plan = plan.Trim();

            cl = new CLineLife(name, policyNum, type, plan, accDate, dueDate, mopd, premium, cRate, split, comm);
            return cl;
        } 
    }
}