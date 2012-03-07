using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Data;

namespace ReportGen
{
    class ExReport
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        List<ProjectReport> projectReportList = new List<ProjectReport>();
        DataTable a;
        private void CreateTables()
        {

        }
        public void read(String filename)
        {
            object misValue = System.Reflection.Missing.Value;
            DateTime _reportDate;
            String _reporterName;

            _reportDate = getDateFromFileName(filename);

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filename, 0, true, 5, "", "", true, 
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            _reporterName = readProjectReport(_reportDate);
            readNonProjectReport(_reporterName, _reportDate);
            xlWorkBook.Close(true, misValue, misValue);

            xlApp.Quit();
            releaseObj(xlWorkSheet);
            releaseObj(xlWorkBook);
            releaseObj(xlApp);
        }

        private void readNonProjectReport(string _reporterName, DateTime _reportDate)
        {
            Excel.Worksheet ws = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            Excel.Range range = ws.UsedRange;
            Debug.Print(range.Row.ToString());
            Object[,] data;
            data = (System.Object[,])range.get_Value(Type.Missing);
            long iRows;
            long iCols;
            iRows = data.GetUpperBound(0);
            iCols = data.GetUpperBound(1);
            Debug.Print(data[6, 2].ToString());
            List<NonProjectReport> reports = new List<NonProjectReport>();
            
            for ( int i = 6 ; i < iRows; i++)
            {
                NonProjectReport report = new NonProjectReport();
                if (data[i, 1] != null)
                {
                    if (data[i, NonProjectReport.cCategory] != null)
                    {
                        report.category = data[i, NonProjectReport.cCategory] == null ? "" :data[i, NonProjectReport.cCategory].ToString().Trim();
                        report.project = data[i, NonProjectReport.cProject] == null ? "" : data[i, NonProjectReport.cProject].ToString().Trim();
                        report.customer = data[i, NonProjectReport.cCustomer] == null ? "" : data[i, NonProjectReport.cCustomer].ToString().Trim();
                        report.report = data[i, NonProjectReport.cReport] == null ? "" : data[i, NonProjectReport.cReport].ToString().Trim();
                        report.personInCharge = data[i, NonProjectReport.cPersonInCharge] == null ? "" : data[i, NonProjectReport.cPersonInCharge].ToString().Trim();
                        report.team = data[i, NonProjectReport.cTeam] == null ? "" : data[i, NonProjectReport.cTeam].ToString().Trim();
                        report.issue = data[i, NonProjectReport.cIssue] == null ? "" : data[i, NonProjectReport.cIssue].ToString().Trim();
                        
                    }
                }


                Debug.Print(data[i, 1].ToString());
                /*
                if (r.Value2 != null)
                {
                    strTemp = r.Value2.ToString();
                    Debug.Print(strTemp);
                }
                */
            }
        }

        private String readProjectReport(DateTime reportDate)
        {
            String _reporterName;
            String _project = null, _record = null, _issue = null, _plan = null;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Excel.Range range = xlWorkSheet.UsedRange;
            
            _reporterName = xlWorkSheet.get_Range("I6").Value2;

            for (int i = 1; i <= range.Rows.Count; i++)
            {
                String cell = "B" + i;
                Excel.Range r;
                ProjectReport item;
                r = xlWorkSheet.get_Range(cell, cell);
                String strTemp = r.Value2;
                if (strTemp != null)
                {
                    strTemp = strTemp.Replace(" ", "");
                }
                if (strTemp == "프로젝트")
                {
                    _project = range[r.Row, r.Column + 1].Value2;
                }
                else if (strTemp == "추진실적")
                {
                    _record = range[r.Row, r.Column + 1].Value2;
                }
                else if (strTemp == "추진계획")
                {
                    _plan = range[r.Row, r.Column + 1].Value2;
                }
                else if (strTemp == "주요이슈사항")
                {
                    _issue = range[r.Row, r.Column + 1].Value2;
                    item = new ProjectReport();
                    item.project = _project;
                    item.record = _record;
                    item.plan = _plan;
                    item.issue = _issue;
                    item.reporter = _reporterName;
                    item.date = reportDate;
                    projectReportList.Add(item);
                    _project = _record = _plan = _issue = null;
                }
            }
            foreach (ProjectReport d in projectReportList)
            {
                Debug.IndentSize = 4;
                Debug.Print("Project : " + d.project);
                Debug.Indent();
                Debug.Print("record : ");
                Debug.Print(d.record);
                Debug.Print("Plan:");
                Debug.Print(d.plan);
                Debug.Print("issue:");
                Debug.Print(d.issue);
                Debug.Unindent();
             
            }
            return _reporterName;
        }


        public DateTime getDateFromFileName(string filename)
        {
            DateTime day;
            String name;
            day = DateTime.Now;
            name = System.IO.Path.GetFileName(filename);
            char[] delimit = new char[] { '_', '.' };
            //string s10 = "TCS-H002_SYS1G2T_주간 업무 보고서(박준형)_110114.xls";
            String[] parsedName = name.Split(delimit);
            day = DateTime.ParseExact(parsedName[3], "yyMMdd", null);
            Debug.Print(parsedName[3]);
            return day;
        }

        public void close()
        {
        }

        private void releaseObj(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Debug.Print("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
