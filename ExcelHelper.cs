using Microsoft.Office.Interop.Excel;
using System;
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
        public void Cell_SetValue(int row,int colunm, object value)//设置改单元格值
        {
            worksheet.Cells[row, colunm]= value;
        }
        public void Cell_SetFormula(int row, int colunm, string formula)//设置单元格计算公式
        {
            worksheet.Cells[row, colunm] = formula;
        }
        public void Cell_Merge(int row_min, int colunm_min, int row_max, int colunm_max, object value)//合并单元格
        {
            worksheet.Cells[row_min, colunm_min] = value;//先将值复制给左上单元格
            //屏蔽系统弹出的询问窗口
            app.DisplayAlerts = false;
            app.AlertBeforeOverwriting = false;
            ((Range)worksheet.Range[worksheet.Cells[row_min, colunm_min], worksheet.Cells[row_max, colunm_max]]).Merge();//将一个区域合并
        }
        public void Cell_SetRowHeight(int row, object value)//设置单元格行高
        {
            try
            {
                ((Range)worksheet.Rows[row]).RowHeight = value;
            }
            catch//如果row超范围，则将所有的行高都调整为目标高度
            {
                ((Range)worksheet.Columns[1]).RowHeight = value;
            }
            
        }
        public void Cell_SetColumnWidth(int colunm, object value)//设置单元格行高
        {
            try
            {
                ((Range)worksheet.Columns[colunm]).ColumnWidth = value;
            }
            catch//如果colunm超范围，则将所有的列宽都调整为目标宽度
            {
                ((Range)worksheet.Rows[1]).ColumnWidth = value;
            }
            
        }
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
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            app = null;
        }
    }
}
