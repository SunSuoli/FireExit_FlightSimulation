using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace FireExit_FlightSimulation
{
    class Excel_NPIO
    {
        FileStream fileStream;//文件流
        IWorkbook workbook = null;//工作簿
        ISheet sheet;  //工作表
        public void Open_Create_WorkBook(String FilePath)
        {
            if (File.Exists(FilePath))//文件存在
            {
                fileStream = new FileStream(FilePath, FileMode.Open, FileAccess.Read);
                if (FilePath.IndexOf(".xlsx") > 0) // 2007版本
                {
                    workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook
                }
                else if (FilePath.IndexOf(".xls") > 0) // 2003版本
                {
                    workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook
                }
                else
                {
                    MessageBox.Show("请选择EXCEL文件！", "提示"); 
                }
            }
            else//文件不存在
            {
                String Extension = Path.GetExtension(FilePath);//文件后缀
                if (Extension == ".xls")
                {
                    workbook = new HSSFWorkbook();  //新建xls工作簿
                }
                else if(Extension == ".xlsx")
                {
                    workbook = new XSSFWorkbook();  //新建xlsx工作簿
                }
                else
                {
                    FilePath += ".xlsx";
                    workbook = new XSSFWorkbook();  //新建xlsx工作簿
                }
                fileStream = new FileStream(FilePath, FileMode.Create);
                workbook.CreateSheet("Sheet1");  //新建3个Sheet工作表
                workbook.CreateSheet("Sheet2");
                workbook.CreateSheet("Sheet3");
                workbook.Write(fileStream);
            }
                
        }
        public void Open_WorkSheet(int idex)
        {
            sheet = workbook.GetSheetAt(idex);
        }
        public List<List<String>> Read_WorkSheet()
        {
            List < List <String>> Sheet_Data=new List<List<string>>();
            //string[,] Sheet_Data=new string[100,100];
            IRow row;// = sheet.GetRow(0);            //新建当前工作表行数据
            for (int i = 0; i < sheet.LastRowNum; i++)  //对工作表每一行
            {
                row = sheet.GetRow(i);   //row读入第i行数据
                if (row != null)
                {
                    Sheet_Data.Add(new List<string>());//在二维列表中增加一行
                    for (int j = 0; j < row.LastCellNum; j++)  //对工作表每一列
                    {
                        Sheet_Data[i].Add(row.GetCell(j).ToString());//将当前单元格数据增加到列表中
                    }
                }
            }
            return Sheet_Data;
        }
    }
}
