using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace TheExcelEdit
{
    /// <SUMMARY>
    /// Microsoft.Office.Interop.ExcelEdit 的摘要说明
    /// </SUMMARY>
    public class ExcelEdit
    {
        public string mFilename;
        public Application app;
        public Workbooks wbs;
        public Workbook wb;
        public Worksheets wss;
        public Worksheet ws;

        public ExcelEdit()
        {

        }

        public void Create()//创建一个Microsoft.Office.Interop.Excel对象
        {
            app = new Application();
            wbs = app.Workbooks;
            wb = wbs.Add(true);
        }

        public void Open(string FileName)//打开一个Microsoft.Office.Interop.Excel文件
        {
            app = new Application();
            wbs = app.Workbooks;
            wb = wbs.Add(FileName);
            mFilename = FileName;
        }

        public List<string> GetSheetNames()
        {
            List<string> sheetNames = new List<string>();
            foreach (Worksheet sheet in wb.Worksheets)
            {
                string sheetName = sheet.Name;
                if (!string.IsNullOrEmpty(sheetName))
                {
                    sheetNames.Add(sheetName);
                }
            }
            return sheetNames;
        }

        public string GetSheetName(int i)
        {
            Worksheet s = GetSheet(i);
            return s.Name;
        }

        public void SetFileName(string FileName)
        {
            mFilename = FileName;
        }

        public Worksheet GetSheet(string SheetName)
        //获取一个工作表
        {
            Worksheet s = (Worksheet)wb.Worksheets[SheetName];
            return s;
        }

        public Worksheet GetSheet(int i)
        //获取一个工作表
        {
            if((i<= wb.Worksheets.Count) && (i>=1) )
            {
                Worksheet s = (Worksheet)wb.Worksheets.get_Item(i);
                return s;
            }
            
            return (Worksheet)wb.Worksheets.get_Item(1);
        }

        public Worksheet AddSheet(string SheetName)
        //添加一个工作表
        {
            Worksheet s = (Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            s.Name = SheetName;
            return s;
        }

        public void DelSheet(string SheetName)//删除一个工作表
        {
            ((Worksheet)wb.Worksheets[SheetName]).Delete();
        }

        public Worksheet ReNameSheet(string OldSheetName, string NewSheetName)//重命名一个工作表一
        {
            Worksheet s = (Worksheet)wb.Worksheets[OldSheetName];
            s.Name = NewSheetName;
            return s;
        }

        public Worksheet ReNameSheet(Worksheet Sheet, string NewSheetName)//重命名一个工作表二
        {
            Sheet.Name = NewSheetName;
            return Sheet;
        }

        public object[,] GetUsedRangeData(string ws)
        {
            Worksheet s = GetSheet(ws);
            object[,] exceldata = (object[,])s.UsedRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);
            return exceldata;
        }

        public void GetUsedRange(string ws, out int x, out int y)
        {
            x = 0; y = 0;
            Worksheet s = GetSheet(ws);
            x = s.UsedRange.Rows.Count;
            y = s.UsedRange.Columns.Count;
        }

        public object[,] GetRangeData(string ws, int a, int b, int x, int y)
        {
            Worksheet s = GetSheet(ws);
            Range c1 = (Range)s.Cells[a, b];
            Range c2 = (Range)s.Cells[x, y];
            Range rng = s.get_Range(c1, c2);
            object[,] exceldata = (object[,])rng.get_Value(XlRangeValueDataType.xlRangeValueDefault);
            return exceldata;
        }

        public void FindItem(int i, string strtext, out int x, out int y)
        {
            x = 0;y = 0;
            Worksheet s = GetSheet(i);
            Range currentFind = s.UsedRange.Find(strtext, Type.Missing,XlFindLookIn.xlValues, XlLookAt.xlWhole,XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false,
                            Type.Missing, Type.Missing);
            if (currentFind != null)
            {
                x = currentFind.Row;
                y = currentFind.Column;
            }           
        }

        public void FindItem(string ws, string strtext, out int x, out int y)
        {
            x = 0; y = 0;
            Worksheet s = GetSheet(ws);
            Range currentFind = s.UsedRange.Find(strtext, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false,
                            Type.Missing, Type.Missing);
            if (currentFind != null)
            {
                x = currentFind.Row;
                y = currentFind.Column;
            }
        }

        public void SetCellValue(Worksheet ws, int x, int y, object value)
        {
            ws.Cells[x, y] = value;
        }

        public void SetCellValue(string ws, int x, int y, object value)
        //ws：要设值的工作表的名称 X行Y列 value 值
        {
            GetSheet(ws).Cells[x, y] = value;
        }

        public string GetCellText(string sheetName, int rowindex, int columnindex)
        {
            string value = null;
            Worksheet sheet = GetSheet(sheetName);
            value = (string)((Range)sheet.Cells[rowindex, columnindex]).Text;
            return value;
        }

        public string[,] GetRangeText(string ws, int a, int b, int x, int y)
        {
            string[,] value = new string[x-a+1,y-b+1];
            Worksheet sheet = GetSheet(ws);
            //get value in A1
            for(int i=a; i<=x; i++)
            {
                for(int j=b; j<=y; j++)
                {
                    value[i-a,j-b] = (string)((Range)sheet.Cells[i, j]).Text;
                }
            }
            return value;
        }

        public void SetCellProperty(Worksheet ws, int Startx, int Starty, int Endx, int Endy, int size, string name, Constants color, Constants HorizontalAlignment)
        //设置一个单元格的属性   字体，   大小，颜色   ，对齐方式
        {
            name = "宋体";
            size = 12;
            color = Constants.xlAutomatic;
            HorizontalAlignment = Constants.xlRight;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Name = name;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Size = size;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Color = color;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).HorizontalAlignment = HorizontalAlignment;
        }

        public void SetCellProperty(string wsn, int Startx, int Starty, int Endx, int Endy, int size, string name, Constants color, Constants HorizontalAlignment)
        {
            //name = "宋体";
            //size = 12;
            //color = Microsoft.Office.Interop.Excel.Constants.xlAutomatic;
            //HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;

            Worksheet ws = GetSheet(wsn);
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Name = name;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Size = size;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Color = color;

            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).HorizontalAlignment = HorizontalAlignment;
        }

        public void UniteCells(Worksheet ws, int x1, int y1, int x2, int y2)
        //合并单元格
        {
            ws.get_Range(ws.Cells[x1, y1], ws.Cells[x2, y2]).Merge(Type.Missing);
        }

        public void UniteCells(string ws, int x1, int y1, int x2, int y2)
        //合并单元格
        {
            GetSheet(ws).get_Range(GetSheet(ws).Cells[x1, y1], GetSheet(ws).Cells[x2, y2]).Merge(Type.Missing);
        }

        public void InsertTable(System.Data.DataTable dt, string ws, int startX, int startY)
        //将内存中数据表格插入到Microsoft.Office.Interop.Excel指定工作表的指定位置 为在使用模板时控制格式时使用一
        {
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {
                    GetSheet(ws).Cells[startX + i, j + startY] = dt.Rows[i][j].ToString();
                }
            }
        }

        public void InsertTable(System.Data.DataTable dt, Worksheet ws, int startX, int startY)
        //将内存中数据表格插入到Microsoft.Office.Interop.Excel指定工作表的指定位置二
        {
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {
                    ws.Cells[startX + i, j + startY] = dt.Rows[i][j];
                }

            }
        }

        public void AddTable(System.Data.DataTable dt, string ws, int startX, int startY)
        //将内存中数据表格添加到Microsoft.Office.Interop.Excel指定工作表的指定位置一
        {

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {
                    GetSheet(ws).Cells[i + startX, j + startY] = dt.Rows[i][j];
                }
            }
        }

        public void AddTable(System.Data.DataTable dt, Worksheet ws, int startX, int startY)
        //将内存中数据表格添加到Microsoft.Office.Interop.Excel指定工作表的指定位置二
        {
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {
                    ws.Cells[i + startX, j + startY] = dt.Rows[i][j];
                }
            }
        }

        public void InsertPictures(string Filename, string ws)
        //插入图片操作一
        {
            GetSheet(ws).Shapes.AddPicture(Filename, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10, 150, 150);
            //后面的数字表示位置
        }

        public void InsertActiveChart(Microsoft.Office.Interop.Excel.XlChartType ChartType, string ws, int DataSourcesX1, int DataSourcesY1, int DataSourcesX2, int DataSourcesY2, Microsoft.Office.Interop.Excel.XlRowCol ChartDataType)
        //插入图表操作
        {
            ChartDataType = Microsoft.Office.Interop.Excel.XlRowCol.xlColumns;
            wb.Charts.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            {
                wb.ActiveChart.ChartType = ChartType;
                wb.ActiveChart.SetSourceData(GetSheet(ws).get_Range(GetSheet(ws).Cells[DataSourcesX1, DataSourcesY1], GetSheet(ws).Cells[DataSourcesX2, DataSourcesY2]), ChartDataType);
                wb.ActiveChart.Location(XlChartLocation.xlLocationAsObject, ws);
            }
        }

        public bool Save()
        //保存文档
        {
            if (mFilename == "")
            {
                return false;
            }
            else
            {
                try
                {
                    wb.Save();
                    return true;
                }

                catch (Exception ex)
                {
                    return false;
                }
            }
        }

        public bool SaveAs(object FileName)
        //文档另存为
        {
            try
            {
                wb.Saved = true;
                wb.SaveCopyAs(FileName);
                return true;
            }

            catch (Exception ex)
            {
                return false;
            }
        }

 //       public class KeyMyExcelProcess
 //       {
 ////           [DllImport("User32.dll", CharSet = CharSet.Auto)]
 //           //public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
 //           public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
 //           {
 //               try
 //               {
 //                   IntPtr t = new IntPtr(excel.Hwnd);   //得到这个句柄，具体作用是得到这块内存入口
 //                   int k = 0;
 //                   GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k
 //                   System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
 //                   p.Kill();     //关闭进程k
 //               }
 //               catch (System.Exception ex)
 //               {
 //                   throw ex;
 //               }
 //           }
 //       }

        public void Close()
        //关闭一个Microsoft.Office.Interop.Excel对象，销毁对象
        {
            //wb.Save();
            wb.Close(false, Type.Missing, Type.Missing);
            wbs.Close();
            app.Quit();
            wb = null;
            wbs = null;
            app = null;
            GC.Collect();
        }
    }
}

