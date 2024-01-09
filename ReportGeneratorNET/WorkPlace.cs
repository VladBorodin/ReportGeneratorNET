using OfficeOpenXml.Style;
using OfficeOpenXml;
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

namespace ReportGeneratorNET {
    public partial class WorkPlace : Form {
        User u = new User();
        List<User> Users = new List<User>();
        List<Doc> ld = ExcelInput();
        private BindingSource bindingSource1 = new BindingSource();
        public WorkPlace(List<User> users, User user) {
            InitializeComponent();
            Users = users;
            u = user;
            if (user.VIP) {
                VIPstrip.Visible = true;
            }
            if (user.guest) {
                dataGridView1.ReadOnly = true;
                файлToolStripMenuItem.Enabled = false;
                groupBox1.Enabled = false;
            }
            this.Text = $"Генератор отчета - {user.name}"; 
            bindingSource1.DataSource = ld;
            dataGridView1.DataSource = bindingSource1;
            dataGridView1.Columns[0].HeaderText = "Дата получения";
            dataGridView1.Columns[1].HeaderText = "№ извещения";
            dataGridView1.Columns[2].HeaderText = "Основание";
            dataGridView1.Columns[3].HeaderText = "Изделие";
            dataGridView1.Columns[4].HeaderText = "Приоритет";
            dataGridView1.Columns[5].HeaderText = "БД";
            dataGridView1.Columns[6].HeaderText = "Состояние";
            dataGridView1.Columns[7].HeaderText = "В ООМ";
            dataGridView1.Columns[8].HeaderText = "Дата изменения";
            dataGridView1.Columns[9].HeaderText = "Кол-во";
            dataGridView1.Columns[10].HeaderText = "Примечание";
        }
        static List<Doc> ExcelInput() {
            List<Doc> ld = new List<Doc>();
            // file name with .xlsx extension  
            string p_strPath = "source\\report.xlsx";
            if (File.Exists(p_strPath)) {
                // Creating an instance 
                // of ExcelPackage
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage excel = new ExcelPackage(p_strPath)) {
                    //get the first worksheet in the workbook
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];
                    int colCount = worksheet.Dimension.End.Column;  //get Column Count
                    int rowCount = worksheet.Dimension.End.Row;     //get row count

                    for (int row = 2; row <= rowCount; row++) {
                        ld.Add(new Doc {
                            date_receive = DateTime.Parse(worksheet.Cells[row, 1].Value.ToString()),
                            nomber_doc = worksheet.Cells[row, 2].Value.ToString(),
                            category = worksheet.Cells[row, 3].Value.ToString(),
                            product = worksheet.Cells[row, 4].Value.ToString(),
                            priority = worksheet.Cells[row, 5].Value.ToString(),
                            database = worksheet.Cells[row, 6].Value.ToString(),
                            executors = worksheet.Cells[row, 7].Value.ToString().Split(',').ToList<string>(),
                            state = worksheet.Cells[row, 8].Value.ToString(),
                            state_confirm = Boolean.Parse(worksheet.Cells[row, 9].Value.ToString()),
                            state_date = DateTime.Parse(worksheet.Cells[row, 10].Value.ToString()),
                            done_work = worksheet.Cells[row, 11].Value.ToString().Split(',').ToList<string>(),
                            rout = worksheet.Cells[row, 12].Value.ToString().Split(',').ToList<string>(),
                            amount = worksheet.Cells[row, 13].Value.ToString(),
                            note = worksheet.Cells[row, 14].Value.ToString()
                        }); ;
                    }
                    excel.Dispose();
                    //Close Excel package
                }
            } else {
                using (FileStream fs = new FileStream(p_strPath, FileMode.Create, FileAccess.Write, FileShare.None)) {
                    // Creating an instance 
                    // of ExcelPackage 
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    ExcelPackage excel = new ExcelPackage();

                    // name of the sheet 
                    var workSheet = excel.Workbook.Worksheets.Add("Извещения");

                    // setting the properties 
                    // of the work sheet  
                    workSheet.TabColor = System.Drawing.Color.Black;
                    workSheet.DefaultRowHeight = 12;

                    // Setting the properties 
                    // of the first row 
                    workSheet.Row(1).Height = 20;
                    workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Row(1).Style.Font.Bold = true;

                    // Header of the Excel sheet 
                    workSheet.Cells[1, 1].Value = "Дата получения";
                    workSheet.Cells[1, 2].Value = "№ извещения";
                    workSheet.Cells[1, 3].Value = "Основание";
                    workSheet.Cells[1, 4].Value = "Изделие";
                    workSheet.Cells[1, 5].Value = "Приоритет";
                    workSheet.Cells[1, 6].Value = "БД";
                    workSheet.Cells[1, 7].Value = "Исполнители";
                    workSheet.Cells[1, 8].Value = "Состояние";
                    workSheet.Cells[1, 9].Value = "Закрыт";
                    workSheet.Cells[1, 10].Value = "Дата изменения";
                    workSheet.Cells[1, 11].Value = "Завершили работу";
                    workSheet.Cells[1, 12].Value = "Расцеховка";
                    workSheet.Cells[1, 13].Value = "Кол-во";
                    workSheet.Cells[1, 14].Value = "Примечание";
                    excel.Dispose();
                    //Close Excel package
                }
            }
            return ld;
        }
        public static void ExcelOutput(List<Doc> ld) {
            // Creating an instance 
            // of ExcelPackage 
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelPackage excel = new ExcelPackage();

            // name of the sheet 
            var workSheet = excel.Workbook.Worksheets.Add("Извещения");

            // setting the properties 
            // of the work sheet  
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            // Setting the properties 
            // of the first row 
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            // Header of the Excel sheet 
            workSheet.Cells[1, 1].Value = "Дата получения";
            workSheet.Cells[1, 2].Value = "№ извещения";
            workSheet.Cells[1, 3].Value = "Основание";
            workSheet.Cells[1, 4].Value = "Изделие";
            workSheet.Cells[1, 5].Value = "Приоритет";
            workSheet.Cells[1, 6].Value = "БД";
            workSheet.Cells[1, 7].Value = "Исполнители";
            workSheet.Cells[1, 8].Value = "Состояние";
            workSheet.Cells[1, 9].Value = "Закрыт";
            workSheet.Cells[1, 10].Value = "Дата изменения";
            workSheet.Cells[1, 11].Value = "Завершили работу";
            workSheet.Cells[1, 12].Value = "Расцеховка";
            workSheet.Cells[1, 13].Value = "Кол-во";
            workSheet.Cells[1, 14].Value = "Примечание";

            // Inserting the article data into excel 
            // sheet by using the for each loop 
            // As we have values to the first row  
            // we will start with second row 
            int recordIndex = 2;

            foreach (var v in ld) {
                workSheet.Cells[recordIndex, 1].Value = v.date_receive.ToShortDateString();
                workSheet.Cells[recordIndex, 2].Value = v.nomber_doc;
                workSheet.Cells[recordIndex, 3].Value = v.category;
                workSheet.Cells[recordIndex, 4].Value = v.product;
                workSheet.Cells[recordIndex, 5].Value = v.priority;
                workSheet.Cells[recordIndex, 6].Value = v.database;
                for (int i = 0; i < v.executors.Count; i++) {
                    workSheet.Cells[recordIndex, 7].Value += $"{v.executors[i]}";
                    if (i != v.executors.Count - 1) workSheet.Cells[recordIndex, 7].Value += ",";
                }
                if (workSheet.Cells[recordIndex, 7].Value == null) {
                    workSheet.Cells[recordIndex, 7].Value += "-";
                } else if (workSheet.Cells[recordIndex, 7].Value.ToString() == "") {
                    workSheet.Cells[recordIndex, 7].Value += "-";
                }
                workSheet.Cells[recordIndex, 8].Value = v.state;
                workSheet.Cells[recordIndex, 9].Value = v.state_confirm;
                workSheet.Cells[recordIndex, 10].Value = v.state_date.ToShortDateString();
                for (int i = 0; i < v.done_work.Count; i++) {
                    workSheet.Cells[recordIndex, 11].Value += $"{v.done_work[i]}";
                    if (i != v.done_work.Count - 1) workSheet.Cells[recordIndex, 11].Value += ",";
                }
                if (workSheet.Cells[recordIndex, 11].Value == null) {
                    workSheet.Cells[recordIndex, 11].Value += "-";
                } else if (workSheet.Cells[recordIndex, 11].Value.ToString() == "") {
                    workSheet.Cells[recordIndex, 11].Value += "-";
                }
                for (int i = 0; i < v.rout.Count; i++) {
                    workSheet.Cells[recordIndex, 12].Value += $"{v.rout[i]}";
                    if (i != v.rout.Count - 1) workSheet.Cells[recordIndex, 12].Value += ",";
                }
                if (workSheet.Cells[recordIndex, 12].Value == null) {
                    workSheet.Cells[recordIndex, 12].Value += "-";
                } else if (workSheet.Cells[recordIndex, 12].Value.ToString() == "") {
                    workSheet.Cells[recordIndex, 12].Value += "-";
                }
                workSheet.Cells[recordIndex, 13].Value = v.amount;
                workSheet.Cells[recordIndex, 14].Value = v.note;
                recordIndex++;
            }

            // By default, the column width is not  
            // set to auto fit for the content 
            // of the range, so we are using 
            // AutoFit() method here.  
            workSheet.Column(1).AutoFit();
            workSheet.Column(2).AutoFit();
            workSheet.Column(3).AutoFit();
            workSheet.Column(4).AutoFit();
            workSheet.Column(5).AutoFit();
            workSheet.Column(6).AutoFit();
            workSheet.Column(7).AutoFit();
            workSheet.Column(8).AutoFit();
            workSheet.Column(9).AutoFit();
            workSheet.Column(10).AutoFit();
            workSheet.Column(11).AutoFit();
            workSheet.Column(12).AutoFit();
            workSheet.Column(13).AutoFit();
            workSheet.Column(14).AutoFit();

            // file name with .xlsx extension  
            string p_strPath = "source\\report.xlsx";

            if (File.Exists(p_strPath))
                File.Delete(p_strPath);

            // Create excel file on physical disk  
            FileStream objFileStrm = File.Create(p_strPath);
            objFileStrm.Close();

            // Write content to excel file  
            File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
            //Close Excel package 
            excel.Dispose();
        }
        private void button2_Click(object sender, EventArgs e) {
            int col;
            if (dataGridView1.CurrentCell.ColumnIndex >= 0) {
                col = dataGridView1.CurrentCell.ColumnIndex;
            } else {
                col = 1;
            }
            if (col == 0) {
                var ordered_data = from n in ld orderby n.date_receive descending select n;
                if (ld.SequenceEqual(ordered_data)) {
                    ordered_data = from n in ld orderby n.date_receive ascending select n;
                }
                ld = ordered_data.ToList<Doc>();
            } else if(col == 1) {
                var ordered_data = from n in ld orderby n.nomber_doc descending select n;
                if (ld.SequenceEqual(ordered_data)) {
                    ordered_data = from n in ld orderby n.nomber_doc ascending select n;
                }
                ld = ordered_data.ToList<Doc>();
            }
            bindingSource1.DataSource = ld;
            dataGridView1.Refresh();
                //dataGridView1.Sort(dataGridView1.Columns[col], ListSortDirection.Ascending);

        }
            private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e) {
            //dataGridView1.Sort(dataGridView1.Columns[dataGridView1.SelectedCells[0].OwningColumn.HeaderText], ListSortDirection.Ascending);
            //MessageBox.Show($"Text: {dataGridView1.CurrentCell.OwningColumn.HeaderText}");
        }

        private void button1_Click(object sender, EventArgs e) {
            ld.Add(new Doc { date_receive = DateTime.Today });
            dataGridView1.Refresh();
        }
        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e) {
            ExcelOutput(ld);
        }
        private void button6_Click(object sender, EventArgs e) {
            dataGridView1.Rows.Remove(dataGridView1.Rows[dataGridView1.CurrentRow.Index]);
            dataGridView1.Refresh();
        }

        private void button4_Click(object sender, EventArgs e) {
            var dataRow = dataGridView1.CurrentCell.OwningRow;
            Doc doc = new Doc();
            doc.date_receive = DateTime.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            bool there = false;
            foreach(Doc d in ld) {
                if (d.nomber_doc == dataGridView1.CurrentRow.Cells[1].Value.ToString()) {
                    there = true;
                    doc = d; break;
                }
            }
            if (!there) {
                doc.date_receive = DateTime.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                doc.nomber_doc = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                doc.category = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                doc.product = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                doc.priority = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                doc.database = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                doc.state = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                doc.state_confirm = Boolean.Parse(dataGridView1.CurrentRow.Cells[7].Value.ToString());
                doc.state_date = DateTime.Parse(dataGridView1.CurrentRow.Cells[8].Value.ToString());
                doc.amount = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                doc.note = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                ld.Add(doc);
            }
            DocWork DW = new DocWork(ld,doc.nomber_doc, Users, u);
            DW.ShowDialog();
        }

        private void button1_Click_1(object sender, EventArgs e) {
            string num = textBox1.Text;
            var row = dataGridView1.Rows[0];
            for(int i = 0; i < dataGridView1.RowCount - 1; i++) {
                row = dataGridView1.Rows[i];
                if (row.Cells[1].Value.ToString() == num) {
                    dataGridView1.Rows[i].Selected = true;
                }
            }
        }
    }
}
