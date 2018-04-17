using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using Microsoft.Office.Interop.Excel;
using System.Reflection;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string filename;
        string ExcelPath;

        private void buttonAddFile_Click(object sender, EventArgs e)
        {
            
            //openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "文本文件|*.txt";
            this.openFileDialog1.ShowDialog();

            //  openFileDialog1.FileOk()
        }

        private void o(object sender, EventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            
            filename = openFileDialog1.FileName;
            this.dataGridView1.Rows.Add(filename);
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count==0)
            {
                MessageBox.Show("请选择需要删除的文件!");
            }

            else
            {
                foreach (DataGridViewRow DataGridViewRowOne in dataGridView1.SelectedRows)  //集合之间怎么操作? one collection in another collection
                                                                                            // DataGridViewRowOne.SetValues("");
                    this.dataGridView1.Rows.Remove(DataGridViewRowOne);
            }
        }

        private void buttonEmpty_Click(object sender, EventArgs e)
        {


            //for (int i=0;i< dataGridView1.Rows.Count;i++)
            //{
            //    //int j = dataGridView1.Rows.Count;
            //    //dataGridView1.Rows.RemoveAt(0);
                
            //}

            dataGridView1.Rows.Clear();

        }

        private void buttonReplace_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text is null)
            {
                MessageBox.Show("请输入需要替换的字符!");
            }
            else
            {
                string SearchStr = this.textBox1.Text;
            }

            if (this.textBox2.Text is null)
            {
                MessageBox.Show("请输入用于替换的字符!");
            }
            else
            {
                string ReplaceStr = this.textBox2.Text;
            }


            ArrayList ListFileNames = new ArrayList();
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                ListFileNames.Add(dataGridView1.Rows[i].Cells[0].Value.ToString());
            }

           

            TxtReplace(ListFileNames, textBox1.Text, textBox2.Text);


        }

        public void TxtReplace(ArrayList StrArraylist, string SearchStr, string RepalceStr)
        {
            foreach (string  item in StrArraylist)
            {
                string temp = "";

                FileStream fs1 = new FileStream(item.ToString(), FileMode.Open,FileAccess.Read);
                StreamReader sr1 = new StreamReader(fs1);
                temp=sr1.ReadToEnd();
                sr1.Close();
                fs1.Close();


                temp = temp.Replace(SearchStr, RepalceStr);

                FileStream fs2 = new FileStream(item.ToString(), FileMode.Open, FileAccess.Write);
                StreamWriter sw2 = new StreamWriter(fs2);
                sw2.WriteLine(temp);
                sw2.Close();
                fs2.Close();

                MessageBox.Show("替换完成!");

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog2.Filter = "Exlce文件|*.xls|Exlce文件|*.xlsx";
            this.openFileDialog2.ShowDialog();

        }

        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
              ExcelPath = this.openFileDialog2.FileName;

        }

        private void buttonNew_Click(object sender, EventArgs e)
        {

            ArrayList SearchList = new ArrayList();
            ArrayList ReplaceList = new ArrayList();

            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            excel.Visible = false;
            excel.UserControl = true;
            // 以可编辑的形式打开EXCEL文件,第三个和第十个参数指定了是否只读和是否可编辑
            //Workbook wb = excel.Application.Workbooks.Open(ExcelPath, missing, false, missing, missing, missing,
            // missing, missing, missing, true, missing, missing, missing, missing, missing);
            Workbooks wbs = excel.Workbooks;
            Workbook wb = wbs.Add(ExcelPath);

            //取得第一个工作薄
            Worksheet ws = (Worksheet)wb.Worksheets[1];
            //取得总记录行数   (包括标题列)
            int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数
            //int columnsint = mySheet.UsedRange.Cells.Columns.Count;//得到列数
            for (int i = 1; i <= rowsint; i++)
            {
                string tempSearch = ((Range)ws.Cells[i, 1]).Text;
                SearchList.Add(tempSearch);
                string tempReplace = ((Range)ws.Cells[i, 2]).Text;
                ReplaceList.Add(tempReplace);
            }

            ArrayList ListFileNames = new ArrayList();
            for (int i = 0; i < dataGridView1.RowCount-1; i++)
            {
                ListFileNames.Add(dataGridView1.Rows[i].Cells[0].Value.ToString());
            }

            for (int i = 0; i < ListFileNames.Count; i++)
            {
                TxtMultyReplace(ListFileNames[i].ToString(), SearchList, ReplaceList);
                ws.Cells[i+1,3] = "转换完成!";//cells[行,列]
            }

            MessageBox.Show("替换完成!");

            //先删除原来的excel文件避免保存时弹出是否替换原有文件的对话框
            File.Delete(ExcelPath);

            wb.SaveAs(ExcelPath, missing, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                2, missing, missing, missing, missing);

            //wb.Save();
            ws = null;
            wb.Close();
            wbs.Close();
            wb = null;
            wbs = null;
            excel.Quit();
            excel = null;
            

        }

        public void TxtMultyReplace(string path,ArrayList SerchArrayList,ArrayList ReplacearrayList)
        {
            FileStream fileStreamNewTxt = new FileStream(path.Replace(".txt", "_copy.txt"), FileMode.Create, FileAccess.Write);
            FileStream fileStreamOldTxt = new FileStream(path, FileMode.Open, FileAccess.Read);

            StreamWriter sw = new StreamWriter(fileStreamNewTxt);
            StreamReader sr1 = new StreamReader(fileStreamOldTxt);
            string OldTxt = sr1.ReadToEnd();

            if (SerchArrayList.Count!=ReplacearrayList.Count)
            {
                MessageBox.Show("搜索字符和替换字符数量不相等!");
                 
            }

            int cou = SerchArrayList.Count;

            for (int i = 0; i < cou; i++)
            {
                
                string NewTxt = OldTxt.Replace(SerchArrayList[i].ToString(), ReplacearrayList[i].ToString());
                sw.WriteLine(NewTxt);
            }

            sr1.Close();
            sw.Close();
            fileStreamOldTxt.Close();
            fileStreamNewTxt.Close();

        }
        
}  
}
