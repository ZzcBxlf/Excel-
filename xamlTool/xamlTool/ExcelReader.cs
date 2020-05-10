using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;

namespace xamlTool
{
    public class ExcelReader
    {
        private IWorkbook sworkbook;
        public ExcelReader()
        {
            sworkbook= null;
        }
        public Boolean TranslateFunction(string filePath,out string changedName)
        {
            try
            {
                string fileName = null;
                string fileNameChanged = null;
                string extension = Path.GetExtension(filePath);
                fileName = filePath.Substring(0, filePath.LastIndexOf("."));
                fileName += extension;
                if (!System.IO.File.Exists(filePath))
                {
                    changedName = fileName;
                    return false;
                }
                FileStream fs = File.OpenRead(filePath);
                if (extension.Equals(".xls"))
                {
                    //把xls文件中的数据写入wk中
                    sworkbook = new HSSFWorkbook(fs);
                }
                else if (extension.Equals(".xlsx"))
                {
                    //把xlsx文件中的数据写入wk中
                    sworkbook = new XSSFWorkbook(fs);
                }
                else
                {
                    changedName = null;
                    return false;
                }
                fs.Close();

                for (int i = 0; i < sworkbook.NumberOfSheets; i++)  //NumberOfSheets是myxls.xls中总共的表数
                {
                    ISheet sheet = sworkbook.GetSheetAt(i);   //读取当前表数据
                    for (int j = 0; j <= sheet.LastRowNum; j++)  //LastRowNum 是当前表的总行数
                    {
                        IRow row = sheet.GetRow(j);  //读取当前行数据
                        if (row != null)
                        {
                            for (int k = 0; k < row.LastCellNum; k++)  //LastCellNum 是当前行的总列数
                            {
                                ICell cell = row.GetCell(k);  //当前表格
                                cell.SetCellType(CellType.String);
                            }
                        }
                    }
                }
                fileNameChanged = filePath.Substring(0, filePath.LastIndexOf("."));
                fileNameChanged +="_changed" + extension;
                changedName = fileNameChanged.Substring(fileNameChanged.LastIndexOf("\\") + 1, (fileNameChanged.LastIndexOf(".") - fileNameChanged.LastIndexOf("\\") - 1));
                using (FileStream newFile = File.OpenWrite(fileNameChanged)) //打开一个xls文件，如果没有则自行创建，如果存在myxls.xls文件则在创建是不要打开该文件！
                {
                    sworkbook.Write(newFile);   //向打开的这个xls文件中写入mySheet表并保存。
                }
                return true;
            }
            catch(Exception ex)
            {
                changedName = null;
                return false;
            }
        }
    }
}
