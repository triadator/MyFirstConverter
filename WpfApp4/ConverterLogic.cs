using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Diagnostics;

namespace WpfApp4
{
    internal class ConverterLogic
    {
        StreamReader streamReader;
        public ConverterLogic(string filepath)
        {
            //Открытие файла
            try
            {
                                               
                StreamReader streamReader = File.OpenText(filepath);
                this.streamReader = streamReader;
            }
            catch (System.IO.IOException)
            {
                MessageBox.Show("Этот документ уже открыт, закройте его и попытайтесь заново");
                throw;
            }
            /*foreach (var process in Process.GetProcessesByName(filepath))
            {
                process.Kill();
            }
            */
        }
        public void Convert()
        {
            string text = streamReader.ReadToEnd();
            streamReader.Close();
            //Убираем точки с запятыми
            string pattern2 = @";";
            string replacement = ",";
            string result = Regex.Replace(text, pattern2, replacement);
            //Удаляем
            result = Regex.Replace(result, @"Имя,Фамилия,Отчество,Дата Рождения,Телефон,email", String.Empty);
            
            //Убираем подпись года            
            result = Regex.Replace(result, @"г\.", String.Empty);
            //Убираем суммы покупки и даты покупки
            result = Regex.Replace(result, @"\,\d{0,10}\,?\d{1,2}?\,\d{1,2}\s\w+\s20\d{2}", "");            
           
            // Удаляем лишние символы пустых строк 
            result = Regex.Replace(result, @"((\d{12,}|ПРЕЗЕНТ-карта)[\,|\s]+)", "");
            //Делаем тройное имя вначале
            result = Regex.Replace(result, @"(\b\w+\b)\,(\b\w+\b)?\,(\b\w+\b)?\,", "$1 $2 $3,$1,$2,$3,");
            //Меняем месяца
            List<string> monthpatterns = new List<string>() { " января ", " февраля ", " марта ", " апреля ", " мая ", " июня ", " июля ", " августа ", " сентября ", " октября ", " ноября ", " декабря " };
            for (int i = 0; i < monthpatterns.Count; i++)
            {
                string monthpattern = monthpatterns[i];
                string target = ".0" + (i + 1) + ".";
                if (i > 8)
                { target = "." + (i + 1) + ".";}
                Regex regex1 = new Regex(monthpattern);
                result = regex1.Replace(result, target);
            }
            // Добавляем к датам месяца нули
            result = Regex.Replace(result, @"(\b\d{1}\p{P}\d{2}\p{P}\d{4})", "0$1");
           
            //Добавляем запятые перед и после др
            result = Regex.Replace(result, @"(\b\d{2}\p{P}\d{2}\p{P}\d{4})", ",,,,,,,,,,$1,,,,,,,,,,,,,,Cуаре");
            
            //Меняем формат даты
            result = Regex.Replace(result, @"(\b\d{2})\p{P}(\d{2})\p{P}(\d{4})", "$3-$1-$2");
            
            //Меняем местами email,mobile, удаляем  род деятельности            
            result = Regex.Replace(result, @"([\w|\-]+)?(\,\b\d{11})((\,)([\w\.\-]+\@[\w\-]+\.\w{2,3})?)", "$3,$2");

            //Меняем формат телефона
            result = Regex.Replace(result, @"\b8(\d{3})(\d{3})(\d{2})(\d{2})", "+7 $1 $2-$3-$4");
                        //Убрать пробелы
            result = Regex.Replace(result, @"(Cуаре)(\s\,)", "$1,*");
         

            // CОХРАНЕНИЕ ФАЙЛА
            SaveFileDialog sD = new SaveFileDialog();
            sD.Filter = "СSV file (*.csv)|*.csv";
            if (sD.ShowDialog() == true)
            {
                StreamWriter sw = new StreamWriter(File.Create(sD.FileName), Encoding.UTF8);
                sw.WriteLine("Name,Given Name,Additional Name,Family Name,Yomi Name,Given Name Yomi,Additional Name Yomi,Family Name Yomi,Name Prefix,Name Suffix,Initials,Nickname,Short Name,Maiden Name,Birthday,Gender,Location,Billing Information,Directory Server,Mileage,Occupation,Hobby,Sensitivity,Priority,Subject,Notes,Language,Photo,Group Membership,E-mail 1 - Type,E-mail 1 - Value,Phone 1 - Type,Phone 1 - Value,Phone 2 - Type,Phone 2 - Value,Phone 3 - Type,Phone 3 - Value,Address 1 - Type,Address 1 - Formatted,Address 1 - Street,Address 1 - City,Address 1 - PO Box,Address 1 - Region,Address 1 - Postal Code,Address 1 - Country,Address 1 - Extended Address,Organization 1 - Type,Organization 1 - Name,Organization 1 - Yomi Name,Organization 1 - Title,Organization 1 - Department,Organization 1 - Symbol,Organization 1 - Location,Organization 1 - Job Description,Organization 2 - Type,Organization 2 - Name,Organization 2 - Yomi Name,Organization 2 - Title,Organization 2 - Department,Organization 2 - Symbol,Organization 2 - Location,Organization 2 - Job Description,Relation 1 - Type,Relation 1 - Value\t");
                sw.WriteLine(result);
                sw.Close();
            }
        }
    }
}
