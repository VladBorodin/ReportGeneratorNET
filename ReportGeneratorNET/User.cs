using System;
using System.Windows.Forms;

namespace ReportGeneratorNET{
    /// <summary>
    /// Класс пользователь, в котором хранится информация из файла со списком табельных номеров
    /// </summary>
    public class User{
        public bool VIP = false;
        public bool guest = false;
        public int login{get;set;}
        public string name {get;set;}
        public void setUser(User user) {
            VIP = user.VIP;
            name = user.name;
        }
    }
}