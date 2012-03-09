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
        private System.Data.DataSet dataSet;

        private void CreateTables()
        {
            // Instantiate the DataSet variable.
            dataSet = new DataSet();

            MakeProjectTable();
            MakeNonProjectTable();
            MakeRelations();

        }

        private void MakeRelations()
        {
            DataRelation dateRelation;
            dateRelation = dataSet.Relations.Add("DateMatch",
                dataSet.Tables["ProjectTable"].Columns["date"],
                dataSet.Tables["NonProjectTable"].Columns["date"]);
        }

        private DataColumn MakeStringRecord(String name)
        {
            DataColumn column;
            // Create second column.
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = name;
            column.AutoIncrement = false;
            column.Caption = name;
            column.ReadOnly = false;
            column.Unique = false;
            return column;
        }

        private void MakeProjectTable()
        {
            // Create a new DataTable.
            System.Data.DataTable table = new DataTable("ProjectTable");
            // Declare variables for DataColumn and DataRow objects.
            DataColumn column;

            // Create new DataColumn, set DataType, 
            // ColumnName and add to DataTable.    
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "id";
            column.ReadOnly = true;
            column.Unique = false;
            // Add the Column to the DataColumnCollection.
            table.Columns.Add(column);

            // Create second column.
            column = MakeStringRecord("project");
            //column = new DataColumn();
            //column.DataType = System.Type.GetType("System.String");
            //column.ColumnName = "project";
            //column.AutoIncrement = false;
            //column.Caption = "project";
            //column.ReadOnly = false;
            //column.Unique = false;
            // Add the column to the table.
            table.Columns.Add(column);

            // Create second column.
            column = MakeStringRecord("record");

            // Add the column to the table.
            table.Columns.Add(column);

            // Create second column.
            column = MakeStringRecord("plan");

            // Add the column to the table.
            table.Columns.Add(column);

            // Create second column.
            column = MakeStringRecord("issue");

            // Add the column to the table.
            table.Columns.Add(column);

            // Create second column.
            column = MakeStringRecord("reporter");

            // Add the column to the table.
            table.Columns.Add(column);

            // Create second column.
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.DateTime");
            column.ColumnName = "date";
            column.AutoIncrement = false;
            column.Caption = "date";
            column.ReadOnly = false;
            column.Unique = false;
            // Add the column to the table.
            table.Columns.Add(column);

            // Make the ID column the primary key column.
            //DataColumn[] PrimaryKeyColumns = new DataColumn[1];
            //PrimaryKeyColumns[0] = table.Columns["id"];
            //table.PrimaryKey = PrimaryKeyColumns;

            // Add the new DataTable to the DataSet.
            dataSet.Tables.Add(table);

            // Create three new DataRow objects and add 
            // them to the DataTable
            /*
            for (int i = 0; i <= 2; i++)
            {
                row = table.NewRow();
                row["id"] = i;
                row["ParentItem"] = "ParentItem " + i;
                table.Rows.Add(row);
            }
            */
        }
        private void MakeNonProjectTable()
        {
            // Create a new DataTable.
            System.Data.DataTable table = new DataTable("NonProjectTable");
            // Declare variables for DataColumn and DataRow objects.
            DataColumn column;

            // Create new DataColumn, set DataType, 
            // ColumnName and add to DataTable.    
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "id";
            column.ReadOnly = true;
            column.Unique = true;
            // Add the Column to the DataColumnCollection.
            table.Columns.Add(column);

            // Create second column.
            column = MakeStringRecord("category");
            table.Columns.Add(column);

            column = MakeStringRecord("project");
            table.Columns.Add(column);

            column = MakeStringRecord("customer");
            table.Columns.Add(column);

            column = MakeStringRecord("report");
            table.Columns.Add(column);

            column = MakeStringRecord("personInCharge");
            table.Columns.Add(column);

            column = MakeStringRecord("team");
            table.Columns.Add(column);

            column = MakeStringRecord("issue");
            table.Columns.Add(column);

            column = MakeStringRecord("reporter");
            table.Columns.Add(column);

            // Create second column.
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.DateTime");
            column.ColumnName = "date";
            column.AutoIncrement = false;
            column.Caption = "date";
            column.ReadOnly = false;
            column.Unique = false;
            // Add the column to the table.
            table.Columns.Add(column);

            // Add the new DataTable to the DataSet.
            dataSet.Tables.Add(table);
        }

        public void read(String filename)
        {
            object misValue = System.Reflection.Missing.Value;
            DateTime _reportDate;
            String _reporterName;
            CreateTables();
            _reportDate = getDateFromFileName(filename);

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filename, 0, true, 5, "", "", true, 
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            _reporterName = readProjectReport(_reportDate);
            readNonProjectReport(_reporterName, _reportDate);
            xlWorkBook.Close(true, misValue, misValue);
            displayDataSet();
            xlApp.Quit();
            releaseObj(xlWorkSheet);
            releaseObj(xlWorkBook);
            releaseObj(xlApp);
        }

        private void displayDataSet()
        {
            DataTable projectTable = dataSet.Tables["ProjectTable"];
            foreach (DataRow row in projectTable.Rows)
            {
                
            }
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

                        System.Data.DataTable nonProjectTable;
                        nonProjectTable = dataSet.Tables["NonProjectTable"];
                        DataRow row;
                        row = nonProjectTable.NewRow();
                        row["category"] = data[i, NonProjectReport.cCategory] == null ? "" : data[i, NonProjectReport.cCategory].ToString().Trim();
                        row["project"] = data[i, NonProjectReport.cProject] == null ? "" : data[i, NonProjectReport.cProject].ToString().Trim();
                        row["customer"] = data[i, NonProjectReport.cCustomer] == null ? "" : data[i, NonProjectReport.cCustomer].ToString().Trim();
                        row["report"] = data[i, NonProjectReport.cReport] == null ? "" : data[i, NonProjectReport.cReport].ToString().Trim();
                        row["personInCharge"] = data[i, NonProjectReport.cPersonInCharge] == null ? "" : data[i, NonProjectReport.cPersonInCharge].ToString().Trim();
                        row["team"] = data[i, NonProjectReport.cTeam] == null ? "" : data[i, NonProjectReport.cTeam].ToString().Trim();
                        row["issue"] = data[i, NonProjectReport.cIssue] == null ? "" : data[i, NonProjectReport.cIssue].ToString().Trim();
                        row["reporter"] = _reporterName;
                        row["date"] = _reportDate;
                        nonProjectTable.Rows.Add(row);                    
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

                    System.Data.DataTable ProjectTable;
                    ProjectTable = dataSet.Tables["ProjectTable"];
                    DataRow row;
                    row = ProjectTable.NewRow();

                    row["project"] = _project;
                    row["record"] = _record;
                    row["plan"] = _plan;
                    row["issue"] = _issue;
                    row["reporter"] = _reporterName;
                    row["date"] = reportDate;
                    row["id"] = 1;
                    ProjectTable.Rows.Add(row);

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
