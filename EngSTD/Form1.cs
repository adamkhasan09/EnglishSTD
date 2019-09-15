using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MaterialSkin;
using MaterialSkin.Controls;

namespace EngSTD
{
    public partial class Form1 : MaterialForm
    {
        int[] ids;
        int id;
        int label_manager_en;
        int label_manager_ru;
        string BasePath = Application.StartupPath + @"\UserBook.xlsx";
        List<int> save_words_ids = new List<int>();
        List<int> delet_user_ids = new List<int>();
        bool user_dcnry;
        public Form1()
        {
            InitializeComponent();
            var skinManager = MaterialSkinManager.Instance;
            skinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            skinManager.ColorScheme = new ColorScheme(Primary.BlueGrey800,Primary.BlueGrey900,Primary.BlueGrey500,Accent.LightBlue700,TextShade.WHITE);
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            bool validation;
            try
            {
                int a = int.Parse(textBox1.Text);
                int b = int.Parse(textBox2.Text);
                validation = b > a ? true: false;
            }
            catch
            {
                validation = false;
            }
            if(validation == true)
            {
                user_dcnry = false;
                if (comboBox1.SelectedIndex == 0)
                {
                    label_manager_en = 1;
                    label_manager_ru = 2;
                }
                else
                {
                    label_manager_en = 2;
                    label_manager_ru = 1;
                }
                Excel excel = new Excel(BasePath, 1);
                Handler HND = new Handler();
                int firstEl, lastEs, diapZn;
                firstEl = int.Parse(textBox1.Text);
                lastEs = int.Parse(textBox2.Text);
                diapZn = Math.Abs(firstEl - lastEs) + 1;
                ids = new int[diapZn];
                for (int i = 0; i < diapZn; i++)
                {
                    ids[i] = firstEl;
                    firstEl += 1;
                }
                id = HND.getRndByArray(ids);
                label1.Text = excel.ReadCell(ids[id], label_manager_en);
                label2.Text = excel.ReadCell(ids[id], label_manager_ru);
                excel.Close();
                label2.Visible = false;
            }
            else
            {
                MessageBox.Show("Пожалуйста, введите корректные данные");
            }
            
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            Excel excel = new Excel(BasePath, 1);
            Handler HND = new Handler();
            if (ids.Length != 1)
            {
                if(user_dcnry == true)
                {
                    delet_user_ids.Add(ids[id]);
                }
                ids = HND.IntArrayReductionByIndex(ids, id);
                id = HND.getRndByArray(ids);
                label1.Text = excel.ReadCell(ids[id], label_manager_en);
                label2.Text = excel.ReadCell(ids[id], label_manager_ru);
                label2.Visible = false;
                
            }
            else
            {
                MessageBox.Show("Дипазон выбранных слов был пройден");
            }
            excel.Close();
        }

        

        private void Button2_Click(object sender, EventArgs e)
        {
            
            Excel excel = new Excel(BasePath, 1);
            Handler HND = new Handler();
            int place_id = HND.getRndByArray(ids);
            int save_val = ids[id];
            save_words_ids.Add(save_val);
            ids[id] = ids[place_id];
            ids[place_id] = save_val;
            label1.Text = excel.ReadCell(ids[id], label_manager_en);
            label2.Text = excel.ReadCell(ids[id], label_manager_ru);
            label2.Visible = false;
            excel.Close();
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            Excel excel = new Excel(BasePath, 2);
            Handler HND = new Handler();
            if (textBox3.Text != "" && user_dcnry == false)
            {
              
                int number = int.Parse(textBox3.Text);
                string str = excel.ReadCell(number, 1);
                if(str == "")
                {
                    save_words_ids = new List<int>(save_words_ids.Distinct());
                    save_words_ids.Sort();
                    string str_ids = HND.intArrayToStr(save_words_ids, ",");
                    excel.WriteCell(int.Parse(textBox3.Text), 1, str_ids);
                }
                else
                {
                    List<int> saved_arr = HND.srtToIntList(str, ',');
                    for (int i = 0; i < save_words_ids.Count(); i++)
                        saved_arr.Add(save_words_ids[i]);
                    saved_arr = new List<int>(saved_arr.Distinct());
                    saved_arr.Sort();
                    string str_ids = HND.intArrayToStr(saved_arr, ",");
                    excel.WriteCell(int.Parse(textBox3.Text), 1, str_ids);
                }
                label8.Visible = false;
                textBox3.Visible = false;
                textBox3.Text = "";
                save_words_ids.Clear();
                excel.Save();
                MessageBox.Show("Слова успешно сохранены");
            }
            else if( textBox3.Text != "" && user_dcnry == true)
            {
                DialogResult result =  MessageBox.Show("Хотите удалить слова которые вы знаете?", 
                    "Уведомление", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result.Equals(DialogResult.Yes))
                {
                    string str = excel.ReadCell(int.Parse(textBox3.Text), 1);
                    List<int> saved_array = HND.srtToIntList(str, ',');
                    delet_user_ids = new List<int>(delet_user_ids.Distinct());
                    for (int i = 0; i < saved_array.Count; i++)
                    {
                        for (int j = 0; j < delet_user_ids.Count; j++)
                        {
                            if (saved_array[i] == delet_user_ids[j])
                            {
                                saved_array.Remove(delet_user_ids[j]);
                            }
                        }
                    }
                    excel.Save();
                    saved_array = new List<int>(saved_array.Distinct());
                    string new_str;
                    if(saved_array.Count > 0)
                    {
                        new_str = HND.intArrayToStr(saved_array, ",");
                        excel.WriteCell(int.Parse(textBox3.Text), 1, new_str);
                        excel.Save();
                        label8.Visible = false;
                        textBox3.Visible = false;
                        textBox3.Text = "";
                        delet_user_ids.Clear();
                        excel.Close();
                        MessageBox.Show("Слова которые вы знаете успешно удалены");
                    }
                }       

            }
            else
            {
                label8.Visible = true;
                textBox3.Visible = true;
            }
            excel.Close();

        }

        private void Button6_Click(object sender, EventArgs e)
        {
            user_dcnry = true;
            bool validation;
            try
            {
                int a = int.Parse(textBox4.Text);
                validation = true;
            }
            catch
            {
                validation = false;
            }
            if(validation == true)
            {
                if (comboBox2.SelectedIndex == 0)
                {
                    label_manager_en = 1;
                    label_manager_ru = 2;
                }
                else
                {
                    label_manager_en = 2;
                    label_manager_ru = 1;
                }
                Excel excel_sheet_1 = new Excel(BasePath, 1);
                Excel excel_sheet_2 = new Excel(BasePath, 2);
                Handler HND = new Handler();
                string str = excel_sheet_2.ReadCell(int.Parse(textBox4.Text), 1);
                List<int> int_list = HND.srtToIntList(str, ',');
                ids = int_list.ToArray();
                id = HND.getRndByArray(ids);
                label1.Text = excel_sheet_1.ReadCell(ids[id], label_manager_en);
                label2.Text = excel_sheet_1.ReadCell(ids[id], label_manager_ru);
                excel_sheet_1.Close();
                delet_user_ids.Clear();
            }
            else
            {
                MessageBox.Show("Пожалуйста, введите корректные данные");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
        }

        private void Button4_Click_1(object sender, EventArgs e)
        {
            label2.Visible = true;
        }
    }
    class Handler
    {
        public int[] IntArrayReductionByIndex(int[] array, int idx)
        {
            int[] newArr = new int[array.Count() -1 ];
            int j = 0;
            for (int i = 0; i < array.Count(); i++)
            {
                if(i != idx)
                {
                    newArr[j] = array[i];
                    j++;
                }
            }
            return newArr;
        }
        public int getRndByArray(int[] array)
        {
            int idx;
            Random rndm = new Random();
            idx = rndm.Next(0, array.Count());
            return idx;
        }
        public string intArrayToStr(List<int>arr, string separator)
        {
            string str ="";
            int len = arr.Count();
            int i = 0;
            foreach (int num in arr)
            {
                if (i != len - 1)
                {
                    str += num.ToString() + separator;
                }
                else
                {
                    str += num.ToString();
                }
                i++;
            }
            return str;
        }
        public List<int> srtToIntList(string srt, char separator)
        {
            List<int> list = new List<int>();
            string[] ids = srt.Split(new char[] {separator});
            for(int i = 0; i < ids.Count(); i++)
            {
                list.Add(int.Parse(ids[i]));
            }
            return list;
        }
    }

}
