using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;

namespace FireExit_FlightSimulation
{
    class ExcelHelper
    {
        private Application app = new Application();//实例化应用
        private Workbooks workbooks;//定义工作簿空间
        private _Workbook workbook ;//定义工作簿
        private Sheets sheets;//定义所有的表
        private _Worksheet worksheet;//定义要操作的表
        private Range range;//定义单元格区域
        public void File_OpenorCreate(String FilePath)
        {
            app = new Application();
            workbooks = app.Workbooks;
            workbook = workbooks.Add(FilePath);
            sheets = workbook.Sheets;
            worksheet = sheets.Item[1];//默认操作第一个表
        }
        public void WorkSheet_Choose(int idex)//选择要操作的表
        {

            worksheet = sheets.Item[idex];
        }
        public void WorkSheet_Rename(string name)//重命名已经选择的表
        {
            worksheet.Name = name;
        }
        public void WorkSheet_Delete()//删除已经选择的表
        {
            app.DisplayAlerts = false;
            worksheet.Delete();
        }
        public void WorkSheet_Add()//在已选择的表之后增加一个表
        {
            app.Worksheets.Add(Missing.Value, worksheet);
        }
        public void Row_Delete(int index)//删除一行
        {
            ((Range)worksheet.Rows[index, Missing.Value]).Delete(XlDeleteShiftDirection.xlShiftUp);
        }
        public void Row_Add(int index)//增加一行
        {
            ((Range)worksheet.Rows[index, Missing.Value]).Insert(Missing.Value, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
        }
        public void Cloum_Delete(int index)//删除一列
        {
            ((Range)worksheet.Columns[index,Missing.Value ]).Delete(XlDeleteShiftDirection.xlShiftToLeft);
        }
        public void Cloum_Add(int index)//增加一列
        {
            ((Range)worksheet.Columns[index,Missing.Value ]).Insert(Missing.Value, XlInsertShiftDirection.xlShiftToRight);
        }
        public void Range_Select(int row_min=-1, int colunm_min=-1, int row_max=-1, int colunm_max=-1)
        {
            if(row_min == -1)//如果行的起始值是-1，则认为从第一行开始
            {
                row_min = 0;
            }
            if (colunm_min == -1)//如果列的起始值是-1.则认为从第一列开始
            {
                colunm_min = 0;
            }
            if (row_max == -1)//如果行的结束值是-1，则认为到最后一行结束
            {
                row_max = worksheet.Rows.Count-1;
            }
            if (colunm_max == -1)//如果列的结束值是-1，则认为到最后一列结束
            {
                colunm_max = worksheet.Columns.Count-1;
            }

            range = (Range)worksheet.Range[worksheet.Cells[row_min+1, colunm_min+1], worksheet.Cells[row_max+1, colunm_max+1]];
        }
        //单元格设置
        public void Range_SetFormula(string formula)//设置单元格计算公式
        {
            range.Value2 = formula;
        }
        public void Range_Merge()//合并单元格
        {
            //屏蔽系统弹出的询问窗口
            app.DisplayAlerts = false;
            app.AlertBeforeOverwriting = false;
            range.Merge();//将一个区域合并
        }
        public void Range_SetRowHeight(double value)//设置单元格行高
        {
            range.RowHeight = value;
        }
        public void Range_SetColumnWidth(double value)//设置单元格行高
        {
            range.ColumnWidth = value;
        }
        public void Range_SetColor(object color)//设置单元格背景颜色,颜色共有56中
        {
           range.Interior.ColorIndex = color;
        }
        public void Range_SetFont_Color(Color color)//设置单元格字体颜色
        {
            range.Font.Color = ColorTranslator.ToOle(color);
        }
        public void Range_SetFont_Size(int size)//设置单元格字体大小
        {
            range.Font.Size = size;
        }
        public void Range_SetFont_Blod(bool bold)//设置单元格字体粗体
        {
            range.Font.Bold = bold;
        }
        public void Range_SetFont_Name(string  name)//设置单元格字体名称
        {
            range.Font.Name = name;
        }

        //单元格写入
        public void Range_SetValue(object[,] value)//设置一个区域值
        {
          range.Value2=value;//设置一个区域的值
        }
        //单元格读取
        public string[,] Range_GetValue()//合并单元格
        {
            string[,] data = new string[range.Rows.Count, range.Columns.Count];
            for(int i=0;i<range.Rows.Count; i++)
            {
                for (int j = 0; j < range.Columns.Count; j++)
                {
                    if(range.Value2[i + 1, j + 1] = null)
                    {
                        data[i, j] = "";
                    }
                    else
                    {
                        data[i, j] = (string)range.Value2[i + 1, j + 1];//EXCEL的索引从1开始
                    }
                }
            }
            return data;
        }

        //文件保存
        public void File_SaveAs(String FilePath)//另存文件
        {
            //屏蔽系统弹出的询问窗口
            app.DisplayAlerts = false;
            app.AlertBeforeOverwriting = false;

            workbook._SaveAs(FilePath);//保存到指定目录
        }
        public void File_Close()//销毁文件引用
        {
            workbook.Close();
            workbooks.Close();
            app.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            //app = null;
        }
    }
}
