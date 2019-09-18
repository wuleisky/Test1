using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using ValidationResult = System.Windows.Controls.ValidationResult;

namespace listbox
{
    /// <summary>
    /// DataError.xaml 的交互逻辑
    /// </summary>
    public partial class DataError : Window
    {
        public DataError()
        {
            InitializeComponent();
            this.DataContext = new Person();
            string d = "1";
            string e = d ?? "ttt";
        }
    }


    class NameExists : ValidationAttribute
    {
        public override bool IsValid(object value)
        {
            var name = value as string;
            // 这里可以到数据库等存储容器中检索
            if (name != "abc")
            {
                return false;
            }
            return true;
        }

        public override string FormatErrorMessage(string name)
        {
            return "请输入存在的用户名。";
        }
    }



    //public class CustomerValidationUtils
    //{
    //    public static System.ComponentModel.DataAnnotations.ValidationResult CheckName(string value)
    //    {
    //        if (value.Length < 8)
    //        {
    //            return new ValidationResult("名字长度必须大于等于8位。");
    //        }
    //        return ValidationResult.Success;
    //    }
    //}



    public class Person : INotifyPropertyChanged, IDataErrorInfo
    {

        private string _name;
        [NameExists]
        public string Name
        {
            get { return _name; }
            set
            {
                if (_name != value)
                {
                    _name = value;
                    RaisePropertyChanged("Name");
                }
            }
        }



        private int _age;
        [Range(19, 99, ErrorMessage = "年龄必须在18岁以上。")]
        public int Age
        {
            get { return _age; }
            set
            {
                if (_age != value)
                {
                    _age = value;
                    RaisePropertyChanged("Age");
                }
            }
        }


        public string Error
        {
            get { return ""; }
        }

        //public string this[string columnName]
        //{
        //    get
        //    {
        //        if (columnName == "Age")
        //        {
        //            if (_age < 18)
        //            {
        //                return "年龄必须在18岁以上。";
        //            }
        //        }
        //        return string.Empty;
        //    }
        //}

        public string this[string columnName]
        {
            get
            {
                var vc = new ValidationContext(this, null, null);
                vc.MemberName = columnName;
                var res = new List<System.ComponentModel.DataAnnotations.ValidationResult>();
                var result = Validator.TryValidateProperty(this.GetType().GetProperty(columnName).GetValue(this, null), vc, res);
                if (res.Count > 0)
                {
                    return string.Join(Environment.NewLine, res.Select(r => r.ErrorMessage).ToArray());
                }
                return string.Empty;
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;

        internal virtual void RaisePropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
   
}
