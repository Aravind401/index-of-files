using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace FDE
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {

        private string _pathname;
        private object misValue;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                var path = Path.Combine(Path.GetFullPath(textBox1.Text));

                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (checkBox1.Checked == false)
                {
                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return;
                    }

                }
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                int i = 1;
                xlWorkSheet.Cells[i, 1] = "NO";
                xlWorkSheet.Cells[i, 2] = "DOCUMENT";
                xlWorkSheet.Cells[i, 3] = "VIDEOS";
                xlWorkSheet.Cells[i, 4] = "IMAGES";
                xlWorkSheet.Cells[i, 5] = "SOFTWARES";
                xlWorkSheet.Cells[i, 6] = "ZIP FILES";
                i++;
                int j = 2, k = 2, l = 2, m = 2, n = 2;
                try
                {
                    var Filescollection = Directory.EnumerateFiles(textBox1.Text, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".rar")
                    || s.EndsWith(".zip") || s.EndsWith(".mkv") || s.EndsWith(".mp4")
                    || s.EndsWith(".jpg") || s.EndsWith(".jpeg") || s.EndsWith(".png") ||
                    s.EndsWith(".pdf") || s.EndsWith(".pptx") || s.EndsWith(".docx") || s.EndsWith(".mkv") || s.EndsWith(".exe"));
                    List<string> collection = Filescollection.ToList();
                    foreach (string file in collection)
                    {
                        xlWorkSheet.Cells[i, 1] = (i - 1).ToString();
                        if (file == null) continue;
                        if (file.EndsWith(".pdf") || file.EndsWith(".pptx") || file.EndsWith(".txt") || file.EndsWith(".docx"))
                        {
                            string[] _filelist = file.Split('\\');
                            foreach (string filelist in _filelist)
                            {
                                if (filelist.EndsWith(".pdf") || filelist.EndsWith(".pptx") || filelist.EndsWith(".txt") || filelist.EndsWith(".docx"))
                                {
                                    xlWorkSheet.Cells[j, 2] = filelist;
                                    i++;
                                    j++;
                                }
                            }
                        }
                        else

                            if (file.EndsWith(".mkv") || file.EndsWith(".mp4"))
                        {
                            string[] _filelist = file.Split('\\');
                            foreach (string filelist in _filelist)
                            {
                                if (filelist.EndsWith(".mkv") || filelist.EndsWith(".mp4"))
                                {
                                    xlWorkSheet.Cells[k, 3] = filelist;
                                    k++;
                                    i++;
                                }
                            }
                        }
                        else if (file.EndsWith(".jpg") || file.EndsWith(".jpeg") || file.EndsWith(".png"))
                        {
                            string[] _filelist = file.Split('\\');
                            foreach (string filelist in _filelist)
                            {
                                if (filelist.EndsWith(".jpg") || filelist.EndsWith(".jpeg") || filelist.EndsWith(".png"))
                                {
                                    xlWorkSheet.Cells[l, 4] = filelist;
                                    l++; i++;
                                }
                            }
                        }
                        else if (file.EndsWith(".exe") || file.EndsWith(".apk"))
                        {
                            string[] _filelist = file.Split('\\');
                            foreach (string filelist in _filelist)
                            {
                                if (filelist.EndsWith(".exe") || filelist.EndsWith(".apk"))
                                {
                                    xlWorkSheet.Cells[m, 5] = filelist;
                                    m++; i++;
                                }
                            }
                        }
                        else if (file.EndsWith(".zip") || file.EndsWith(".rar"))
                        {
                            string[] _filelist = file.Split('\\');
                            foreach (string filelist in _filelist)
                            {
                                if (filelist.EndsWith(".zip") || filelist.EndsWith(".rar"))
                                {
                                    xlWorkSheet.Cells[n, 6] = filelist;
                                    n++; i++;
                                }
                            }
                        }
                    }
                    //application
                    if (checkBox1.Checked == true)
                    {
                        if (File.Exists(path))
                        {
                            TaskDialog.Show(" File All Ready Exist");
                            return;
                        }
                        var fs = new FileStream(path + "\\" + "Details.txt", FileMode.OpenOrCreate, FileAccess.Write);
                        var sw = new StreamWriter(fs);
                        int inc = 1;
                        foreach (string filelist in collection)
                        {
                            if (filelist == null) continue;
                            sw.Write(inc + "." + filelist.ToString() + Environment.NewLine);
                            inc++;

                        }
                        TaskDialog.Show(" Completed", "");
                    }
                    if (checkBox1.Checked == false)
                    {
                        xlWorkSheet.Range["A1", "A" + i].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        xlWorkSheet.Columns.AutoFit();
                        xlWorkSheet.Rows.AutoFit();
                        xlWorkBook.SaveAs(path + "\\" + "Details.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlWorkBook.Close(true, misValue, misValue);
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlWorkSheet);
                        Marshal.ReleaseComObject(xlWorkBook);
                        Marshal.ReleaseComObject(xlApp);

                        TaskDialog.Show("Completed");
                    }
                }
                catch (Exception ex) { TaskDialog.Show("" + ex); }
            }
            catch (Exception ex) { TaskDialog.Show("select the location"); }



        }
        private void fileextract(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                textBox1.Text = dialog.FileName;
                _pathname = dialog.FileName;
                textBox1.Text = _pathname;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {

            TaskDialog.Show("Task completed");
        }

        private void label2_Click(object sender, EventArgs e)
        {
            Form2 frm = new Form2();
            frm.ShowDialog();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
