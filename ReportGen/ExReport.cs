using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace ReportGen
{
    class ExReport
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        List<ExData> projectReportList = new List<ExData>();

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

            xlWorkBook.Close(true, misValue, misValue);

            xlApp.Quit();
            releaseObj(xlWorkSheet);
            releaseObj(xlWorkBook);
            releaseObj(xlApp);
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
                ExData item;
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
                    item = new ExData();
                    item.project = _project;
                    item.record = _record;
                    item.plan = _plan;
                    item.issue = _issue;
                    item.reporterName = _reporterName;
                    item.reportDate = reportDate;
                    projectReportList.Add(item);
                    _project = _record = _plan = _issue = null;
                }
            }
            foreach (ExData d in projectReportList)
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
