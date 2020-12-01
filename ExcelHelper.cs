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
        public void Row_Delete(int index)//删除指定的行
        {
            ((Range)worksheet.Rows[index, Missing.Value]).Delete(XlDeleteShiftDirection.xlShiftUp);
        }
        public void Row_Add(int index)//在指定的行增加一行
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
        public void File_SaveAs(String FilePath)//另存文件
        {
            app.AlertBeforeOverwriting = false;//屏蔽掉系统跳出的Alert
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
