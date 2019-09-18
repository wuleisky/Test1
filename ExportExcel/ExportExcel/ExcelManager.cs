using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;


namespace ExportExcel
{
    public class CourseCodeInfo
    {
        public string Name { get; set; }
    }

    public class ExcelManager
    {
        public static void SetCellDropdownList(ISheet sheet, int firstcol, int lastcol, string[] vals)
        {
            ////设置生成下拉框的行和列 
            //var cellRegions = new CellRangeAddressList(1, 65535, firstcol, lastcol);

            ////设置 下拉框内容
            //DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(vals);

            ////绑定下拉框和作用区域，并设置错误提示信息
            //HSSFDataValidation dataValidate = new HSSFDataValidation(cellRegions, constraint);
            //dataValidate.CreateErrorBox("输入不合法", "请输入或选择下拉列表中的值。");
            //dataValidate.ShowPromptBox = true;

            //sheet.AddValidationData(dataValidate);

            //HSSFWorkbook workbook = new HSSFWorkbook();
            //HSSFSheet sheet = workbook.createSheet("Data Validation");
            //CellRangeAddressList addressList = new CellRangeAddressList(
            //    0, 0, 0, 0);
            //DVConstraint dvConstraint = DVConstraint.createExplicitListConstraint(
            //    new String[] { "10", "20", "30" });
            //DataValidation dataValidation = new HSSFDataValidation
            //    (addressList, dvConstraint);
            //dataValidation.setSuppressDropDownArrow(false);
            //sheet.addValidationData(dataValidation);
        }

        // <summary>
        /// The add validation.
        /// </summary>
        /// <param name="sheet">
        /// 要加入列表的sheet
        /// </param>
        /// <param name="itemSheet">
        /// 选项 sheet.
        /// </param>
        /// <param name="headerCell">
        /// 标题单元格
        /// </param>
        /// <param name="items">
        /// 列表项
        /// </param>
        private static void AddValidation(ISheet sheet, ISheet itemSheet, ICell headerCell, List<string> items)
        {
            // 新建行
            var row = itemSheet.CreateRow(itemSheet.PhysicalNumberOfRows);

            // 新行中写入选项
            for (int i = 0; i < items.Count; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(items[i]);
            }

            // 要加下拉列表的范围
            var addressList = new CellRangeAddressList(
                headerCell.RowIndex + 1,
                65535,
                headerCell.ColumnIndex,
                headerCell.ColumnIndex);

            var dvHelper = sheet.GetDataValidationHelper();

            // 格式 Sheet2!$A$1:$E$1
            var dvConstraint = dvHelper.CreateFormulaListConstraint(
                $"{itemSheet.SheetName}!$A${row.RowNum + 1}:${(items.Count)}${row.RowNum + 1}");
            var validation = dvHelper.CreateValidation(dvConstraint, addressList);

            // 强制必须填下拉列表给出的值            
            // validation.ShowErrorBox = true;

            sheet.AddValidationData(validation);
        }

        public static void test1()
        {
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();
            HSSFSheet sheet2 = hssfworkbook.CreateSheet("ShtDictionary") as HSSFSheet;
            sheet2.CreateRow(0).CreateCell(0).SetCellValue("itemA");
            sheet2.CreateRow(1).CreateCell(0).SetCellValue("itemB");
            sheet2.CreateRow(2).CreateCell(0).SetCellValue("itemC");
         

            HSSFName range = hssfworkbook.CreateName() as HSSFName;
            range.RefersToFormula = "ShtDictionary!$A1:$A3";
            range.NameName = "dicRange";
          

            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1") as HSSFSheet;
            CellRangeAddressList regions = new CellRangeAddressList(0, 65535, 0, 0);

            DVConstraint constraint = DVConstraint.CreateFormulaListConstraint("dicRange");
            HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
            sheet1.AddValidationData(dataValidate);
            MemoryStream ms=new MemoryStream();
            hssfworkbook.Write(ms);
            string workbookFile = @"D:\\wulei1.xls";
            hssfworkbook = null;
            FileStream fs = new FileStream(workbookFile, FileMode.Create, FileAccess.Write);
            byte[] data = ms.ToArray();
            fs.Write(data, 0, data.Length);
            fs.Flush();
            fs.Close();
        }


        /// 获取Excel列名

        /// </summary>

        /// <param name="columnNumber">列的序号</param>

        /// <returns></returns>

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

        public static void test4()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("sheet1");

            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("姓名");
            row.CreateCell(1).SetCellValue("身份证号");
            row.CreateCell(2).SetCellValue("年级");
            row.CreateCell(3).SetCellValue("班级");
            row.CreateCell(4).SetCellValue("课程");
            row.CreateCell(5).SetCellValue("角色（班主任、单科老师）");


            IRow row1 = sheet.CreateRow(1);
            row1.CreateCell(0).SetCellValue("张峰");
            row1.CreateCell(1).SetCellValue("1111111111111");
            row1.CreateCell(2).SetCellValue("小学六年级");
            row1.CreateCell(3).SetCellValue("4班");
            row1.CreateCell(4).SetCellValue("语文");
            row1.CreateCell(5).SetCellValue("单科老师");

            var ic = workbook.CreateCellStyle();
            ic.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");

            sheet.SetDefaultColumnStyle(1, ic);

            sheet.SetColumnWidth(1, 5000);
            sheet.SetColumnWidth(2, 4000);
            sheet.SetColumnWidth(3, 4000);
            sheet.SetColumnWidth(5, 24000);

            List<CourseCodeInfo> list = new List<CourseCodeInfo>();
            list.Add(new CourseCodeInfo(){ Name = "语文"});
            list.Add(new CourseCodeInfo() { Name = "数学" });
            list.Add(new CourseCodeInfo() { Name = "英议事" });
            var CourseSheetName = "Course";
            var RangeName = "dicRange";
            ISheet CourseSheet = workbook.CreateSheet(CourseSheetName);
            CourseSheet.CreateRow(0).CreateCell(0).SetCellValue("课程列表（用于生成课程下拉框，请勿修改）");
            for (var i = 1; i < list.Count; i++)
            {
                CourseSheet.CreateRow(i).CreateCell(0).SetCellValue(list[i - 1].Name);
            }

            IName range = workbook.CreateName();
            range.RefersToFormula = string.Format("{0}!$A$2:$A${1}", CourseSheetName, list.Count.ToString());
            range.NameName = RangeName;
            //
            CellRangeAddressList regions = new CellRangeAddressList(1, 65535, 4, 4);
            DVConstraint constraint = DVConstraint.CreateFormulaListConstraint(RangeName);
            HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
            sheet.AddValidationData(dataValidate);

            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;
            string workbookFile = @"D:\\777.xls";
          //  sheet2 = null;

            workbook = null;
            FileStream fs = new FileStream(workbookFile, FileMode.Create, FileAccess.Write);
            byte[] data = ms.ToArray();
            fs.Write(data, 0, data.Length);
            fs.Flush();
            fs.Close();
        }

        public static void test3()
        {
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();
            HSSFSheet sheet2 = hssfworkbook.CreateSheet("ShtDictionary") as HSSFSheet;
            HSSFRow dataRow = sheet2.CreateRow(0) as HSSFRow;
            dataRow.CreateCell(0).SetCellValue("省份");
            dataRow.CreateCell(1).SetCellValue("湖北");
            dataRow.CreateCell(2).SetCellValue("湖南");
            dataRow.CreateCell(3).SetCellValue("广东");
           
            dataRow = sheet2.CreateRow(1) as HSSFRow;
            dataRow.CreateCell(0).SetCellValue("湖北");
            dataRow.CreateCell(1).SetCellValue("汉口");
            dataRow.CreateCell(2).SetCellValue("汉阳");
            dataRow.CreateCell(3).SetCellValue("武昌");
            dataRow = sheet2.CreateRow(2) as HSSFRow;
            dataRow.CreateCell(0).SetCellValue("湖南");
            dataRow.CreateCell(1).SetCellValue("长沙");
            dataRow.CreateCell(2).SetCellValue("岳阳");
            dataRow.CreateCell(3).SetCellValue("长沙南");
            dataRow = sheet2.CreateRow(3) as HSSFRow;
            dataRow.CreateCell(0).SetCellValue("广东");
            dataRow.CreateCell(1).SetCellValue("深圳");
            dataRow.CreateCell(2).SetCellValue("广州");
            dataRow.CreateCell(3).SetCellValue("广州东");

            //  sheet2.IsRightToLeft = false;
            IName range1 = hssfworkbook.CreateName();//创建名称
            range1.NameName = "省份";//设置名称
                                   // var colName = GetExcelColumnName(colIndex);//根据序号获取列名，具体代码见下文
                                   range1.RefersToFormula = "ShtDictionary!$B1:D1";

                                   range1 = hssfworkbook.CreateName();//创建名称
                                   range1.NameName = "湖北";//设置名称
                                   // var colName = GetExcelColumnName(colIndex);//根据序号获取列名，具体代码见下文
                                   range1.RefersToFormula = "ShtDictionary!$B2:D2";

                                   range1 = hssfworkbook.CreateName();//创建名称
                                   range1.NameName = "湖南";//设置名称
                                   // var colName = GetExcelColumnName(colIndex);//根据序号获取列名，具体代码见下文
                                   range1.RefersToFormula = "ShtDictionary!$B3:D3";

                                   range1 = hssfworkbook.CreateName();//创建名称
                                   range1.NameName = "广东";//设置名称
                                   // var colName = GetExcelColumnName(colIndex);//根据序号获取列名，具体代码见下文
                                   range1.RefersToFormula = "ShtDictionary!$B4:D4";
            //range1.RefersToFormula = string.Format("{0}!${3}${2}:${3}${1}",
            //    "ShtDictionary",
            //    "4",
            //    2,
            //    "A");

            // var colName = GetExcelColumnName(1);

            HSSFName range = hssfworkbook.CreateName() as HSSFName;
          //  range.RefersToFormula = "ShtDictionary!$B1:D1";
            range.RefersToFormula = "ShtDictionary!$A1:A4";
            range.NameName = "dicRange";


            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1") as HSSFSheet;
            CellRangeAddressList regions = new CellRangeAddressList(0, 0, 0, 0);

            DVConstraint constraint = DVConstraint.CreateFormulaListConstraint("dicRange");
            HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
            sheet1.AddValidationData(dataValidate);

            regions = new CellRangeAddressList(0, 0, 1, 1);

            DVConstraint constraint1 = DVConstraint.CreateFormulaListConstraint(string.Format("INDIRECT(${0}${1})", "A", 1));
          //  constraint = DVConstraint.CreateFormulaListConstraint("dicRange");
            dataValidate = new HSSFDataValidation(regions, constraint1);
            sheet1.AddValidationData(dataValidate);

            //regions = new CellRangeAddressList(2, 2, 0, 0);

            // constraint1 = DVConstraint.CreateFormulaListConstraint(string.Format("INDIRECT(${0}${1})", "C", 2));
            ////  constraint = DVConstraint.CreateFormulaListConstraint("dicRange");
            //dataValidate = new HSSFDataValidation(regions, constraint1);
            //sheet1.AddValidationData(dataValidate);


            MemoryStream ms = new MemoryStream();
            hssfworkbook.Write(ms);
            ms.Flush();
            ms.Position = 0;
            string workbookFile = @"D:\\8888.xls";
            sheet2 = null;
            
            hssfworkbook = null;
            FileStream fs = new FileStream(workbookFile, FileMode.Create, FileAccess.Write);
            byte[] data = ms.ToArray();
            fs.Write(data, 0, data.Length);
            fs.Flush();
            fs.Close();
        }

        public static void test2()
        {
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();
            HSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1") as HSSFSheet;
            CellRangeAddressList regions = new CellRangeAddressList(0, 65535, 0, 0);
            DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(new string[] { "itemA111", "itemB22", "itemC33" });
            HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
            sheet1.AddValidationData(dataValidate);
            MemoryStream ms = new MemoryStream();
            hssfworkbook.Write(ms);
            string workbookFile = @"D:\\wulei1111.xls";
            hssfworkbook = null;
            FileStream fs = new FileStream(workbookFile, FileMode.Create, FileAccess.Write);
            byte[] data = ms.ToArray();
            fs.Write(data, 0, data.Length);
            fs.Flush();
            fs.Close();
        }

        //public static void test1()
        //{
        //    HSSFWorkbook hssfworkbook = new HSSFWorkbook();
        //    HSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1") as HSSFSheet;
        //    CellRangeAddressList regions = new CellRangeAddressList(0, 65535, 0, 0);
        //    DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(new string[] { "itemA", "itemB", "itemC" });
        //    HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
        //    sheet1.AddValidationData(dataValidate);
        //    MemoryStream ms = new MemoryStream();
        //    hssfworkbook.Write(ms);
        //    string workbookFile = @"D:\\wulei22.xls";
        //    hssfworkbook = null;
        //    FileStream fs = new FileStream(workbookFile, FileMode.Create, FileAccess.Write);
        //    byte[] data = ms.ToArray();
        //    fs.Write(data, 0, data.Length);
        //    fs.Flush();
        //    fs.Close();
        //}

        public static void setdownlist()
        {
            //创建工作簿
            HSSFWorkbook ssfworkbook = new HSSFWorkbook();
            //创建工作表(页)
            HSSFSheet sheet1 = ssfworkbook.CreateSheet("Sheet1") as HSSFSheet;
            //创建一行
            HSSFRow headerRow = (HSSFRow)sheet1.CreateRow(0);
            //设置表头
            headerRow.CreateCell(0).SetCellValue("ID");
            //设置表头的宽度 
            sheet1.SetColumnWidth(0, 15 * 256);
            #region     添加显示下拉列表
            HSSFSheet sheet2 = ssfworkbook.CreateSheet("ShtDictionary") as HSSFSheet;
            ssfworkbook.SetSheetHidden(1, true);//隐藏
            sheet2.CreateRow(0).CreateCell(0).SetCellValue("itemA");//列数据
            sheet2.CreateRow(1).CreateCell(0).SetCellValue("itemB");
            sheet2.CreateRow(2).CreateCell(0).SetCellValue("itemC");
            HSSFName range = ssfworkbook.CreateName() as HSSFName;//创建名称
           // range.Reference = "ShtDictionary!$A$1:$A$3";//格式
            range.NameName = "dicRange";
            #endregion
            headerRow.CreateCell(1).SetCellValue("Selected");
            sheet1.SetColumnWidth(1, 15 * 256);
            //将下拉列表添加
            CellRangeAddressList regions = new CellRangeAddressList(1, 65535, 1, 1);
            DVConstraint constraint = DVConstraint.CreateFormulaListConstraint("dicRange");
            HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
            sheet1.AddValidationData(dataValidate);

            headerRow.CreateCell(2).SetCellValue("VALUE");
            sheet1.SetColumnWidth(2, 15 * 256);

            //写入数据
            //创建数据行
            HSSFRow dataRow = (HSSFRow)sheet1.CreateRow(1);
            //填充数据
            dataRow.CreateCell(0).SetCellValue("1");//id
            dataRow.CreateCell(1).SetCellValue("");//选择框
            dataRow.CreateCell(2).SetCellValue("值");//选择框
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            ssfworkbook.Write(ms);
            string filename = "Sheet1" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString() + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + ".xls";
            object Response = null;
            string workbookFile = @"D:\\wulei.xls";
         
            FileStream fs = new FileStream(workbookFile, FileMode.Create, FileAccess.Write);
            byte[] data = ms.ToArray();
            fs.Write(data, 0, data.Length);
            fs.Flush();
            fs.Close();
            //Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + filename + ""));
            //Response.BinaryWrite(ms.ToArray());
            ms.Close();
            ms.Dispose();
        }


    }

    public class ExcelManager<T>
    {
    /// <summary>  
        /// 导出Excel  
        /// </summary>  
        /// <param name="lists"></param>  
        /// <param name="head">中文列名对照</param>  
        /// <param name="workbookFile">保存路径</param>  
        public static void getExcel(List<T> lists, Hashtable head, string workbookFile)
        {
            try
            {
                HSSFWorkbook workbook = new HSSFWorkbook();
                MemoryStream ms = new MemoryStream();
                HSSFSheet sheet = workbook.CreateSheet() as HSSFSheet;
                HSSFRow headerRow = sheet.CreateRow(0) as HSSFRow;
                bool h = false;
                int j = 1;
                Type type = typeof(T);
                PropertyInfo[] properties = type.GetProperties();

                foreach (T item in lists)
                {
                    HSSFRow dataRow = sheet.CreateRow(j) as HSSFRow;
                    int i = 0;
                    foreach (PropertyInfo column in properties)
                    {
                        if (!h)
                        {
                            headerRow.CreateCell(i).SetCellValue(head[column.Name] == null ? column.Name : head[column.Name].ToString());
                            dataRow.CreateCell(i).SetCellValue(column.GetValue(item, null) == null ? "" : column.GetValue(item, null).ToString());
                        }
                        else
                        {
                            dataRow.CreateCell(i).SetCellValue(column.GetValue(item, null) == null ? "" : column.GetValue(item, null).ToString());
                        }

                        i++;
                    }
                    h = true;
                    j++;
                }
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                sheet = null;
                headerRow = null;
                workbook = null;
                FileStream fs = new FileStream(workbookFile, FileMode.Create, FileAccess.Write);
                byte[] data = ms.ToArray();
                fs.Write(data, 0, data.Length);
                fs.Flush();
                fs.Close();
                data = null;
                ms = null;
                fs = null;
            }
            catch (Exception ee)
            {
                string see = ee.Message;
            }
        }
        ///// <summary>  
        ///// 导入Excel  
        ///// </summary>  
        ///// <param name="lists"></param>  
        ///// <param name="head">中文列名对照</param>  
        ///// <param name="workbookFile">Excel所在路径</param>  
        ///// <returns></returns>  
        //public List<T> fromExcel(Hashtable head, string workbookFile)
        //{
        //    try
        //    {
        //        HSSFWorkbook hssfworkbook;
        //        List<T> lists = new List<T>();
        //        using (FileStream file = new FileStream(workbookFile, FileMode.Open, FileAccess.Read))
        //        {
        //            hssfworkbook = new HSSFWorkbook(file);
        //        }
        //        HSSFSheet sheet = hssfworkbook.GetSheetAt(0) as HSSFSheet;
        //        IEnumerator rows = sheet.GetRowEnumerator();
        //        HSSFRow headerRow = sheet.GetRow(0) as HSSFRow;
        //        int cellCount = headerRow.LastCellNum;
        //        //Type type = typeof(T);  
        //        PropertyInfo[] properties;
        //        T t = default(T);
        //        for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
        //        {
        //            HSSFRow row = sheet.GetRow(i) as HSSFRow;
        //            t = Activator.CreateInstance<T>();
        //            properties = t.GetType().GetProperties();
        //            foreach (PropertyInfo column in properties)
        //            {
        //                int j = headerRow.Cells.FindIndex(delegate (Cell c)
        //                {
        //                    return c.StringCellValue == (head[column.Name] == null ? column.Name : head[column.Name].ToString());
        //                });
        //                if (j >= 0 && row.GetCell(j) != null)
        //                {
        //                    object value = valueType(column.PropertyType, row.GetCell(j).ToString());
        //                    column.SetValue(t, value, null);
        //                }
        //            }
        //            lists.Add(t);
        //        }
        //        return lists;
        //    }
        //    catch (Exception ee)
        //    {
        //        string see = ee.Message;
        //        return null;
        //    }
        //}
        object valueType(Type t, string value)
        {
            object o = null;
            string strt = "String";
            if (t.Name == "Nullable`1")
            {
                strt = t.GetGenericArguments()[0].Name;
            }
            switch (strt)
            {
                case "Decimal":
                    o = decimal.Parse(value);
                    break;
                case "Int":
                    o = int.Parse(value);
                    break;
                case "Float":
                    o = float.Parse(value);
                    break;
                case "DateTime":
                    o = DateTime.Parse(value);
                    break;
                default:
                    o = value;
                    break;
            }
            return o;
        }

     
    }
}
