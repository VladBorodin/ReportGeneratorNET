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
using System.Xml.Serialization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Windows.Media.Effects;
using System.Security.Cryptography.Xml;

namespace ReportGeneratorNET{
    public partial class Form1 : Form{
        public string path_file_CSV = "source\\list\\list_of_employees.csv";
        public string path_dir_CSV = "source\\list";
        public Form1(){
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e){
            WriteCSVFile(path_dir_CSV, path_file_CSV);
            List<User> users = readCSVFile(path_file_CSV);
            User user = new User();
            try {
                user.login = int.Parse(textBox1.Text.ToString());
                user = userSelect(users, user.login);
                WorkPlace workPlace = new WorkPlace(users, user);
                MessageBox.Show($"Добро пожаловать в систему {user.name}!");
                this.Hide();
                workPlace.ShowDialog();
                Close();
            } catch (Exception ex) {
                MessageBox.Show($"Введены некорректные символы!\nЛог ошибки: {ex.Message}");
                textBox1.Clear();
                Show();
            }
        }
        /// <summary>
        /// Создание и перезапись списка табельных номеров
        /// </summary>
        /// <param name="path_dir_CSV">Путь к папке со списком табельных номеров</param>
        /// <param name="path_file_CSV">Путь к файлу со списком табельных номеров</param>
        static void WriteCSVFile(string path_dir_CSV, string path_file_CSV){
            if (!Directory.Exists(path_dir_CSV))
            {
                Directory.CreateDirectory(path_dir_CSV);
            }
            if (!File.Exists(path_file_CSV))
            {
                using (FileStream fs = new FileStream(path_file_CSV, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    string outData = "Бисерова И.Г." + ";" + "2518" + ";" + true + "\n" +
                                        "Бородин В.А." + ";" + "5837" + ";" + false + "\n" +
                                        "Бубнов В.Б." + ";" + "7186" + ";" + false + "\n" +
                                        "Воробьева Е.А." + ";" + "0626" + ";" + false + "\n" +
                                        "Данодина М.А." + ";" + "2529" + ";" + false + "\n" +
                                        "Казаев В.А." + ";" + "4677" + ";" + false + "\n" +
                                        "Князева А.А." + ";" + "1649" + ";" + false + "\n" +
                                        "Кондрат В.В." + ";" + "8761" + ";" + false + "\n" +
                                        "Ледовских Н.Г." + ";" + "8160" + ";" + false + "\n" +
                                        "Макарова М.В." + ";" + "2116" + ";" + false + "\n" +
                                        "Маслякова А.А." + ";" + "9444" + ";" + false + "\n" +
                                        "Оберт В.Е" + ";" + "3073" + ";" + true + "\n" +
                                        "Пискунова Г.А." + ";" + "6294" + ";" + false + "\n" +
                                        "Пугина Н.В." + ";" + "6984" + ";" + false + "\n" +
                                        "Пчелинцева А.Я." + ";" + "1011" + ";" + false + "\n" +
                                        "Рассейкина О.В." + ";" + "7949" + ";" + false + "\n" +
                                        "Самсонова Е.А." + ";" + "3180" + ";" + false;
                    byte[] data = Encoding.Default.GetBytes(outData);
                    fs.Write(data, 0, data.Length);
                }
            }
        }
        static List<User> readCSVFile(string path_file_CSV){
            List<User> users = new List<User>();
            using (FileStream fs = new FileStream(path_file_CSV, FileMode.Open, FileAccess.Read, FileShare.Read)){
                byte[] inData = new byte[fs.Length];                                                            //хардкор - дали весь вес файла
                fs.Read(inData, 0, inData.Length);
                string usr = Encoding.Default.GetString(inData);
                string[] userData = usr.Split('\n');
                foreach (string u in userData)
                {
                    if (u.Length == 0) break;
                    string[] user = u.Split(';');
                    users.Add(new User
                    {
                        name = user[0],
                        login = Int32.Parse(user[1]),
                        VIP = bool.Parse(user[2]),
                    });
                }
            }
            return users;
        }
        static User userSelect(List<User> users, int ID){
            User user_return = new User();
            try
            {
                foreach (var u in users)
                {
                    if (u.login == ID)
                    {
                        user_return = u;
                        break;
                    }
                }
                if (user_return.login == 0)
                {
                    MessageBox.Show("Вашего табельного нет в списке!\nПроверьте правильно ли его ввели или работайте в гостевом режиме!\nЕсли номер введен корректно обратитесь к начальнику, чтобы вас добавили в список!");
                    User Guest = new User() { name = "Гость", VIP = false, login = 0, guest = true };
                    user_return = Guest;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show($"Похоже файл со списком сотрудников открыт другим пользователем!\nЛог ошибки:{e.Message}");
            }
            return user_return;
        }
        private void button2_Click(object sender, EventArgs e) {
            this.Close();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Enter) {
                button1_Click(sender, e);
            }
        }
    }
}

