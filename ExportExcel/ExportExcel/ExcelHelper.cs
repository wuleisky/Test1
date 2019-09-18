using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportExcel
{
    public class Category
    {
        public string Name { get; set; }
        public string FatherName { get; set; }
    }

    public class ExcelHelper
    {
        public static void CreateDropDownList(List<Category> lists,string fatherName,string fileName)
        {
            if (lists.Count > 0)
            {
                HSSFWorkbook hssfworkbook = new HSSFWorkbook();
                HSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1") as HSSFSheet;
                CellRangeAddressList regions = new CellRangeAddressList(0, 0, 0, 0);

                HSSFSheet sheet2 = hssfworkbook.CreateSheet("ShtDictionary") as HSSFSheet;
                List<Category> BigCategory = lists.Where(f => f.FatherName == fatherName).ToList();
                int row = 0;
                int column = 1;
                HSSFRow dataRow = sheet2.CreateRow(row++) as HSSFRow;

                IName range1 = hssfworkbook.CreateName();//创建名称
                range1.NameName = fatherName;//设置名称
                 var colName = GetExcelColumnName(BigCategory.Count+1);//根据序号获取列名，具体代码见下文
                range1.RefersToFormula = string.Format("ShtDictionary!$B1:{0}1",colName);

                foreach (var Category in BigCategory)
                {
                    dataRow.CreateCell(0).SetCellValue(fatherName);
                    dataRow.CreateCell(column++).SetCellValue(Category.Name);
                    HSSFRow childrenrow = sheet2.CreateRow(row++) as HSSFRow;
                    childrenrow.CreateCell(0).SetCellValue(Category.Name);
                    List<Category> childrenCategory = lists.Where(f => f.FatherName == Category.Name).ToList();
                    int childcolumn = 1;
                    if (childrenCategory.Count > 0)
                    {
                        foreach (var ca in childrenCategory)
                        {
                            childrenrow.CreateCell(childcolumn++).SetCellValue(ca.Name);
                        }

                         range1 = hssfworkbook.CreateName();//创建名称
                        range1.NameName = Category.Name;//设置名称
                         colName = GetExcelColumnName(childrenCategory.Count + 1);//根据序号获取列名，具体代码见下文
                        range1.RefersToFormula = string.Format("ShtDictionary!$B{1}:{0}{1}", colName,row);
                    }
                 
                }

                

                DVConstraint constraint = DVConstraint.CreateFormulaListConstraint(fatherName);
                HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
                sheet1.AddValidationData(dataValidate);

                regions = new CellRangeAddressList(0, 0, 1, 1);

                DVConstraint constraint1 = DVConstraint.CreateFormulaListConstraint(string.Format("INDIRECT(${0}${1})", "A", 1));
                //  constraint = DVConstraint.CreateFormulaListConstraint("dicRange");
                dataValidate = new HSSFDataValidation(regions, constraint1);
                sheet1.AddValidationData(dataValidate);

                MemoryStream ms = new MemoryStream();
                hssfworkbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                string workbookFile = fileName;
                sheet2 = null;

                hssfworkbook = null;
                FileStream fs = new FileStream(workbookFile, FileMode.Create, FileAccess.Write);
                byte[] data = ms.ToArray();
                fs.Write(data, 0, data.Length);
                fs.Flush();
                fs.Close();
            }
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;
            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }
    }
}
