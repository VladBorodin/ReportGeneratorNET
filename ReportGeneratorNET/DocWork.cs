using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReportGeneratorNET {
    public partial class DocWork : Form {
        private BindingSource bs1 = new BindingSource();
        private BindingSource bs2 = new BindingSource();
        private BindingSource bs3 = new BindingSource();
        private BindingSource bsc1 = new BindingSource();
        List<User> us = new List<User>();
        List<Doc> ldoc = new List<Doc>();
        User u = new User();
        Doc doc = new Doc();
        public DocWork(List<Doc> ld, string num, List<User> users, User user) {
            InitializeComponent();
            ldoc = ld;
            foreach (Doc d in ld) {
                if (d.nomber_doc == num) {
                    doc = d;
                    break;
                }
            }
            u = user;
            us = users;
            this.Text = $"Изв {doc.nomber_doc} - {u.name}";
            List<string> names = new List<string>();
            bool there = false;
            foreach(var n in users) {
                there = false;
                foreach(var l in doc.executors) {
                    if(n.name == l) {
                        there = true;
                        break;
                    }
                }
                if(!there) {
                    names.Add(n.name);
                }
            }
            if (u.VIP == false) { 
                comboBox1.Enabled = false; 
            }
            bsc1.DataSource = names;
            comboBox1.DataSource = bsc1;
            bs1.DataSource = doc.executors;
            listBox1.DataSource = bs1;
            listBox2.DataSource = doc.done_work;
            groupBox2.Text = $"Информация по {doc.nomber_doc}";
            listBox3.DataSource = doc.rout;
            label3.Text = $"Открыто изв №{doc.nomber_doc}";
            label4.Text = $"Дата начала обработки {doc.date_receive.ToShortDateString()}";
            label5.Text = $"Дата изменения {doc.state_date.ToShortDateString()}";
            label8.Text = $"Состояние изв. {doc.state}";
            if (doc.executors.Contains(u.name)) {
                label7.Text = $"Вы, {u.name}, являетесь исполнителем";
            } else {
                label7.Text = $"Вы, {u.name}, НЕ являетесь исполнителем";
            }
        }
        private void button2_Click(object sender, EventArgs e) {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e) {
            bool there = false;
            foreach(var n in doc.executors) {
                if (u.name == n) there = true;
            }
            if (!there) {
                doc.executors.Add(u.name);
            }
            this.Update();
            this.Refresh();
            listBox1.Update();
            listBox1.Refresh();
            listBox2.Update();
            listBox2.Refresh();
            comboBox1.Update();
            comboBox1.Refresh();
            check_line(doc.executors);
        }

        private void comboBox1_Click(object sender, EventArgs e) {
            bool there = false;
            foreach (var n in doc.executors) {
                if (u.name == n) there = true;
            }
            if (!there) doc.executors.Add(u.name);
            foreach (var n in us) {
                if (comboBox1.Items[comboBox1.SelectedIndex].ToString() == n.name) {
                    doc.executors.Add(n.name);
                    listBox1.Items.Clear();
                }
            }
            listBox1.ClearSelected();
            listBox1.Update();
            listBox1.Refresh();
            listBox2.Update();
            listBox2.Refresh();
            comboBox1.Update();
            comboBox1.Refresh();
        }
        public void check_line(List<string> check) {
            if (check.Count == 0) {
                check.Add("-");
            }
            if (check.Count>1 && check.Contains("-")) {
                check.Remove("-");
            }
        }
        private void button4_Click(object sender, EventArgs e) {
            if (listBox1.SelectedItems.Count > 0 && listBox2.SelectedItems.Count == 0) {
                doc.done_work.Add(listBox1.Items[listBox1.SelectedIndex].ToString());
                doc.executors.Remove(listBox1.Items[listBox1.SelectedIndex].ToString());
                //listBox2.Items.Add(listBox1.SelectedIndex);
                //listBox1.Items.Remove(listBox1.SelectedIndex);
                check_line(doc.done_work);
                check_line(doc.executors);
            } else if(listBox1.SelectedItems.Count==0 && listBox2.SelectedItems.Count>0) {
                doc.executors.Add(listBox2.Items[listBox2.SelectedIndex].ToString());
                doc.done_work.Remove(listBox2.Items[listBox2.SelectedIndex].ToString());
                //listBox1.Items.Add(listBox2.SelectedIndex);
                //listBox2.Items.Remove(listBox2.SelectedIndex);
                check_line(doc.executors);
                check_line(doc.done_work);
            } else {
                MessageBox.Show("Выделите фамилию для переноса");
            }
        }

        private void listBox2_Click(object sender, EventArgs e) {
            listBox1.ClearSelected();
        }

        private void listBox1_Click(object sender, EventArgs e) {
            listBox2.ClearSelected();
        }

        private void button5_Click(object sender, EventArgs e) {
            if (listBox1.SelectedItems.Count > 0 && listBox2.SelectedItems.Count == 0) {
                doc.executors.Remove(listBox1.Items[listBox1.SelectedIndex].ToString());
                check_line(doc.done_work);
                check_line(doc.executors);
            } else if (listBox1.SelectedItems.Count == 0 && listBox2.SelectedItems.Count > 0) {
                doc.done_work.Remove(listBox2.Items[listBox2.SelectedIndex].ToString());
                check_line(doc.executors);
                check_line(doc.done_work);
            } else {
                MessageBox.Show("Выделите фамилию для удаления");
            }
        }

        private void button6_Click(object sender, EventArgs e) {
            string str = textBox1.Text;
            if (str.Length==0) {
                MessageBox.Show("Заполните поле названием или номером цеха!");
            } else {
                if (str.Contains(' ')) {
                    str.Replace(" ","_");
                }
                if (str.Length>0) {
                    doc.rout.Add(str);
                }
            }
            check_line(doc.rout);
        }

        private void button7_Click(object sender, EventArgs e) {
            doc.rout.Remove(listBox3.SelectedItem.ToString());
            check_line(doc.rout);
        }

        private void button2_Click_1(object sender, EventArgs e) {
            this.Close();
        }
    }
}
