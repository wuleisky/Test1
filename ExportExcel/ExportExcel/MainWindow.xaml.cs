
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
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExportExcel
{

    public class People
    {
        public string Name { get; set; }
        public string Age { get; set; }
    }

    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
           List<People> list=new List<People>();
           list.Add(new People(){  Name = "张三",Age = "10"});
           list.Add(new People() { Name = "李四", Age = "11" });
           list.Add(new People() { Name = "王五", Age = "13"});
           list.Add(new People() { Name = "345张三", Age = "17" });
           Hashtable ht=new Hashtable();
            ht.Add("Name", "姓名");
            ht.Add("Age", "年龄");
            string path = "D:\\21.xls";
          ExcelManager<People>.getExcel(list,ht,path);
            
        }


        public static void SetCellDropdownList(ISheet sheet, int firstcol, int lastcol, string[] vals)
        {
            //设置生成下拉框的行和列
            var cellRegions = new CellRangeAddressList(1, 65535, firstcol, lastcol);

            //设置 下拉框内容
            DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(vals);

            //绑定下拉框和作用区域，并设置错误提示信息
            HSSFDataValidation dataValidate = new HSSFDataValidation(cellRegions, constraint);
            dataValidate.CreateErrorBox("输入不合法", "请输入或选择下拉列表中的值。");
            dataValidate.ShowPromptBox = true;

            sheet.AddValidationData(dataValidate);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ExcelManager.test3();
            //   HSSFWorkbook workbook = new HSSFWorkbook();
            //ISheet sheet = workbook.CreateSheet("sheet1");
            //ExcelManager.SetCellDropdownList(sheet, 1, 1, new List<string>() { "男", "女", "保密" }.ToArray());
            //MemoryStream ms=new MemoryStream();
            //workbook.Write(ms);
            //ms.Flush();
            //ms.Position = 0;
            //sheet = null;
            //string workbookFile = @"D:\\wulei.xls";
            //workbook = null;
            //FileStream fs = new FileStream(workbookFile, FileMode.Create, FileAccess.Write);
            //byte[] data = ms.ToArray();
            //fs.Write(data, 0, data.Length);
            //fs.Flush();
            //fs.Close();
        }

        public void Createdownlist()
        {
          //  HSSFWorkbook hssfworkbook = new HSSFWorkbook();
          //  ISheet sheet2 = hssfworkbook.CreateSheet("ShtDictionary");
          //  sheet2.CreateRow(0).CreateCell(0).SetCellValue("itemA");
          //  sheet2.CreateRow(1).CreateCell(0).SetCellValue("itemB");
          //  sheet2.CreateRow(2).CreateCell(0).SetCellValue("itemC");
          //  //然后定义一个名称，指向刚才创建的下拉项的区域：

          //  HSSFName range = hssfworkbook.CreateName();
          //  range.Reference = "ShtDictionary!$A1:$A3";
          //  range.NameName = "dicRange";
          ////  最后，设置数据约束时指向这个名称而不是字符数组：

          //  HSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1");
          //  CellRangeAddressList regions = new CellRangeAddressList(0, 65535, 0, 0);
          //  DVConstraint constraint = DVConstraint.CreateFormulaListConstraint("dicRange");
          //  HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
          //  sheet1.AddValidationData(dataValidate);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
           // ExcelManager.test3();
           List<Category>  list=new List<Category>();
           list.Add(new Category() { Name = "湖南", FatherName = "省份" });
           list.Add(new Category() { Name = "湖北", FatherName = "省份" });
           list.Add(new Category() { Name = "河南", FatherName = "省份" });
           list.Add(new Category() { Name = "长沙", FatherName = "湖南" });
           list.Add(new Category() { Name = "当阳", FatherName = "湖南" });
           list.Add(new Category() { Name = "光谷", FatherName = "湖北" });
           list.Add(new Category() { Name = "汉口", FatherName = "湖北" });
           list.Add(new Category() { Name = "汉阳", FatherName = "湖北" });
           list.Add(new Category() { Name = "信阳1", FatherName = "河南" });
           list.Add(new Category() { Name = "信阳2", FatherName = "河南" });
           list.Add(new Category() { Name = "信阳3", FatherName = "河南" });
           list.Add(new Category() { Name = "信阳4", FatherName = "河南" });
            ExcelHelper.CreateDropDownList(list,"省份","D:\\999.xls");
        }
        //public static void SetCellDropdownList(ISheet sheet, int firstcol, int lastcol, string[] vals)
        //{
        //    //设置生成下拉框的行和列
        //    var cellRegions = new CellRangeAddressList(1, 65535, firstcol, lastcol);

        //    //设置 下拉框内容
        //    DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(vals);

        //    //绑定下拉框和作用区域，并设置错误提示信息
        //    HSSFDataValidation dataValidate = new HSSFDataValidation(cellRegions, constraint);
        //    dataValidate.CreateErrorBox("输入不合法", "请输入或选择下拉列表中的值。");
        //    dataValidate.ShowPromptBox = true;

        //    sheet.AddValidationData(dataValidate);

        //}
    }
}
