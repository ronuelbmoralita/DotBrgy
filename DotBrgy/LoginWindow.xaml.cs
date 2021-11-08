using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;
using Path = System.IO.Path;

namespace DotBrgy
{
    /// <summary>
    /// Interaction logic for LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {
        //sqlite connection

        readonly string sqliteConnectionString = @"Data Source=C:\DotBrgy\Database\dotBrgyDatabase.sqlite;Version=3;";
        public DispatcherTimer timer = new DispatcherTimer();   

        //

        public LoginWindow()
        {
            InitializeComponent();
            timer.Interval = TimeSpan.FromSeconds(1);
            timer.Tick += Timer_Tick;
        }

        /////////////////////////////////////////////////////////////////////////////TIMER

        private int i = 30;

        private void Timer_Tick(object sender, EventArgs e)
        {
            i--;
            count.Text = "DotBrgy is locked. Try again in" + " " + i.ToString() + " s.";
            count.Visibility = Visibility.Visible;
            if (i == 0)
            {
                count.Visibility = Visibility.Collapsed;
                this.IsEnabled = true;
                timer.Stop();
                i = 30;
                Normal_cursor();
            }
        }

        //

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            userName.Focus();
            CreateFolder();
            CreateTable();
        }

        //

        private void CreateFolder()
        {
            string root = @"C:\DotBrgy";
            //string hiddenRoot = @"c:\DotBrgy\Database\Backup"; // or whatever 
            //string subdir = @"C:\Temp\Mahesh";
            // If directory does not exist, create it. 
            if (Directory.Exists(root))
            {
                return;
            }
            else
            {
                Directory.CreateDirectory(root);
                Directory.CreateDirectory(Path.Combine(root, @"Images"));
                Directory.CreateDirectory(Path.Combine(root, @"Documents"));
                Directory.CreateDirectory(Path.Combine(root, @"Database"));
                Directory.CreateDirectory(Path.Combine(root, @"Backup\Database")).Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                Directory.CreateDirectory(Path.Combine(root, @"Print"));
                //DirectoryInfo di = Directory.CreateDirectory(hiddenRoot);
                //di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
            }
        }

        //

        private void CreateTable()
        {
            if (File.Exists(@"C:\DotBrgy\Database\dotBrgyDatabase.sqlite"))
            {
                return;
            }
            else
            {
                SQLiteConnection.CreateFile(@"C:\DotBrgy\Database\dotBrgyDatabase.sqlite");

                string sql = @"CREATE TABLE dotBrgyData(
                                NOS INTEGER PRIMARY KEY AUTOINCREMENT ,
                                residentCode             TEXT      NULL,
                                dateRegister             TEXT      NULL,
                                firstname                TEXT      NULL,  
                                middlename               TEXT      NULL,
                                lastname                 TEXT      NULL,
                                purok                    TEXT      NULL,
                                CivilStatus                   TEXT      NULL,
                                Sex                   TEXT      NULL,
                                Birthdate                 TEXT      NULL,
                                age                      TEXT      NULL,
                                email                    TEXT      NULL
                            );

                                CREATE TABLE userAccount(
	                               NOS INTEGER PRIMARY KEY AUTOINCREMENT ,
                                   userType           TEXT NULL,
                                   username            TEXT NULL,
                                   password            TEXT NULL,
                                   firstname           TEXT NULL,
                                   lastname            TEXT NULL,
                                   address             TEXT NULL
                            );

                                CREATE TABLE history(
	                               NOS INTEGER PRIMARY KEY AUTOINCREMENT ,
                                   residentCode             TEXT      NULL,
                                   dateTime                 TEXT      NULL,
                                   firstname                TEXT      NULL,  
                                   middlename               TEXT      NULL,
                                   lastname                 TEXT      NULL,
                                   documents                TEXT      NULL,
                                   purok                    TEXT      NULL,
                                   Sex                   TEXT      NULL,
                                   Birthdate                 TEXT      NULL,
                                   email                    TEXT      NULL
                            );

                                CREATE TABLE vaccination(
                                NOS INTEGER PRIMARY KEY AUTOINCREMENT ,
                                residentCode             TEXT      NULL,
                                dateTime                 TEXT      NULL,
                                choice                   TEXT      NULL,
                                firstname                TEXT      NULL,  
                                middlename               TEXT      NULL,
                                lastname                 TEXT      NULL,
                                purok                    TEXT      NULL,
                                CivilStatus                   TEXT      NULL,
                                Sex                   TEXT      NULL,
                                Birthdate                 TEXT      NULL,
                                age                      TEXT      NULL,
                                email                    TEXT      NULL
                            );

                                CREATE TABLE statistic(
	                               NOS INTEGER PRIMARY KEY AUTOINCREMENT ,
                                   purok             TEXT      NULL,
                                   residents         TEXT      NULL
                            );

                                CREATE TABLE trialBegin(
	                               NOS INTEGER PRIMARY KEY AUTOINCREMENT ,
                                   dateBegin         TEXT      NULL         
                            );

                                CREATE TABLE trialEnd(
	                               NOS INTEGER PRIMARY KEY AUTOINCREMENT ,
                                   dateEnd         TEXT      NULL         
                            );";

                using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                {
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, con))
                    {
                        con.Open();
                        cmd.ExecuteNonQuery();
                        InsertUser();
                        InsertTrial();
                    }
                }
            }
        }

        //

        private void InsertUser()
        {
            using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
            {
                using (SQLiteCommand cmd = con.CreateCommand())
                {
                    con.Open();
                    cmd.CommandType = CommandType.Text;
                    //cmd.CommandText = "insert into userAccount(userType, username, password, firstname, lastname, address) values(@userType ,@username,@password,@firstname,@lastname,@address)";
                    cmd.CommandText = "insert into userAccount(userType, username, password, firstname, lastname, address) values(@userType ,@username,@password,@firstname,@lastname,@address)";
                    cmd.Parameters.AddWithValue("@userType", "Administrator");
                    cmd.Parameters.AddWithValue("@username", "a");
                    cmd.Parameters.AddWithValue("@password", "a");
                    cmd.Parameters.AddWithValue("@firstname", "Dot");
                    cmd.Parameters.AddWithValue("@lastname", "Brgy");
                    cmd.Parameters.AddWithValue("@address", "cawayan");

                    using (SQLiteCommand command = new SQLiteCommand("Select count (*) from userAccount where username = 'dotBrgy2021'", con))
                    {
                        using (SQLiteCommand command1 = new SQLiteCommand("Select count (*) from userAccount where password = 'brgy2021'", con))
                        {
                            var result = command.ExecuteScalar();
                            int i = Convert.ToInt32(result);
                            var result1 = command.ExecuteScalar();
                            int i1 = Convert.ToInt32(result1);
                            if (i != 0 || i1 != 0)
                            {
                                return;
                            }
                            else
                            {
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
        }

        //

        private void InsertTrial()
        {
            using (SQLiteConnection conBegin = new SQLiteConnection(sqliteConnectionString))
            {
                using (SQLiteCommand cmdBegin = conBegin.CreateCommand())
                {

                    conBegin.Open();
                    cmdBegin.CommandText = "insert into trialBegin(dateBegin) values(@begin)";
                    cmdBegin.Parameters.AddWithValue("@begin", DateTime.Now.ToString());
                    cmdBegin.ExecuteNonQuery();
                    //MessageBox.Show("Working");
                }
            }

            using (SQLiteConnection conEnd = new SQLiteConnection(sqliteConnectionString))
            {
                using (SQLiteCommand cmdEnd = conEnd.CreateCommand())
                {
                    conEnd.Open();
                    cmdEnd.CommandText = "insert into trialEnd(dateEnd) values(@end)";
                    cmdEnd.Parameters.AddWithValue("@end", DateTime.Now.ToString());
                    cmdEnd.ExecuteNonQuery();
                    //MessageBox.Show("Working");
                }
            }
        }

        //clear

        private void Clear(DependencyObject obj)
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {

                if (obj is TextBox textbox)
                    textbox.Text = string.Empty;
                if (obj is CheckBox checkbox)
                    checkbox.IsChecked = false;
                if (obj is ComboBox combobox)
                    combobox.Text = string.Empty;
                if (obj is RadioButton radiobutton)
                    radiobutton.IsChecked = false;
                if (obj is PasswordBox passwordbox)
                    passwordbox.Password = string.Empty;

                Clear(VisualTreeHelper.GetChild(obj, i));
            }
        }

        //mouse cursor

        private void Wait_cursor()
        {
            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait; // set the cursor to loading spinner  
        }

        private void Normal_cursor()
        {
            Mouse.OverrideCursor = null; // set the cursor back to normal
        }

        //

        int attempts = 5;

        private void Login_button_Click(object sender, RoutedEventArgs e)
        {
                ////Regex regex = new Regex("[^a-zA-Z0-9_.]+");
            Regex rx = new Regex(@"^[^a-zA-ZñÑ0-9_.]+");

            if (userName.Text == string.Empty || userPassword.Password == string.Empty)
            {
                MessageBox.Show(this, "Please enter valid username or password!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                Clear(this);
                return;
            }

            else if (rx.IsMatch(userName.Text) || rx.IsMatch(userPassword.Password))
            {
                MessageBox.Show(this, "Not accepting special character's!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                Clear(this);
                return;
            }
            else if (login_checkboxAdmin.IsChecked == false && login_checkboxStandard.IsChecked == false)
            {
                MessageBox.Show(this, "Please select atleast 1 user type!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }
            else if (login_checkboxAdmin.IsChecked == true)
                using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                {
                    using (SQLiteCommand cmd_username = con.CreateCommand())
                    {
                        con.Open();
                        cmd_username.CommandType = CommandType.Text;
                        cmd_username.CommandText = "Select * from userAccount where username = '" + userName.Text.Trim() + "'and password = '" + userPassword.Password.Trim() + "'";
                        SQLiteDataReader sdr_admin;
                        sdr_admin = cmd_username.ExecuteReader();
                        //int count_admin = 0;
                        string userRole_admin = string.Empty;

                        while (sdr_admin.Read())
                        {
                            //count_admin = count_admin + 1;
                            userRole_admin = sdr_admin["userType"].ToString();
                        }

                        if (attempts > 1 && userRole_admin != "Administrator")
                        {
                            attempts--;
                            MessageBox.Show("Invalid Username or Password! " + attempts + " attempts left!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                            Clear(this);
                        }
                        else if (userRole_admin == "Administrator")
                        {
                            //MessageBox.Show(this, "Successfully Login!", "DotBrgy", MessageBoxButtons.OK, MessageBoxIcon.Information);   
                            //username = textbox_username.Text;
                            Wait_cursor();
                            AdminWindow open = new AdminWindow();
                            open.Show();
                            this.Hide();
                            Normal_cursor();
                        }
                        else if (attempts > 0)
                        {
                            MessageBox.Show("Access denied, application will freeze for 30s!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                            this.IsEnabled = false;
                            timer.Start();
                            //MessageBox.Show("Access denied, the application will close!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                            Wait_cursor();
                            attempts = 5;
                            Clear(this);
                            //Application.Current.Shutdown();
                        }
                    }
                }
            else if (login_checkboxStandard.IsChecked == true)
            {
                using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                {
                    using (SQLiteCommand cmd_staff = con.CreateCommand())
                    {
                        con.Open();
                        cmd_staff.CommandType = CommandType.Text;
                        cmd_staff.CommandText = "Select * from userAccount where username = '" + userName.Text.Trim() + "' and password = '" + userPassword.Password.Trim() + "'";

                        SQLiteDataReader sdr_standard;
                        sdr_standard = cmd_staff.ExecuteReader();
                        //int count_standard = 0;
                        string userRole_standard = string.Empty;
                        while (sdr_standard.Read())
                        {
                            //count_standard = count_standard + 1;
                            userRole_standard = sdr_standard["userType"].ToString();
                        }

                        /*
                        if (attempts > 1 && userRole_standard != "Standard")
                        {
                            attempts--;
                            MessageBox.Show(this, "Invalid Username or Password " + attempts + " attempts left!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                            normal_cursor();
                            Clear(this);
                            return;
                        }

                        else if (attempts == 1)
                        {
                            MessageBox.Show(this, "Access denied, the application will close!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Error);
                            normal_cursor();
                            Application.Current.Shutdown();
                        }

                        else if (userRole_standard == "Standard")
                        {
                            //MessageBox.Show(this, "Successfully Login!", "DotBrgy", MessageBoxButtons.OK, MessageBoxIcon.Information);   
                            //username = textbox_username.Text;
                            wait_cursor();
                            MainWindow open = new MainWindow();
                            open.Show();
                            this.Hide();
                            normal_cursor();
                        }   */


                        if (attempts > 1 && userRole_standard != "Standard" && userRole_standard != "Administrator")
                        {
                            attempts--;
                            MessageBox.Show("Invalid Username or Password! " + attempts + " attempts left!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                            Clear(this);
                        }
                        else if (userRole_standard == "Standard" || userRole_standard == "Administrator")
                        {
                            //MessageBox.Show(this, "Successfully Login!", "DotBrgy", MessageBoxButtons.OK, MessageBoxIcon.Information);   
                            //username = textbox_username.Text;
                            Wait_cursor();
                            MainWindow open = new MainWindow();
                            open.Show();
                            this.Hide();
                            Normal_cursor();
                        }
                        else if (attempts > 0)
                        {
                            MessageBox.Show("Access denied, application will freeze for 30s!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                            this.IsEnabled = false;
                            timer.Start();
                            //MessageBox.Show("Access denied, the application will close!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                            Wait_cursor();
                            attempts = 5;
                            Clear(this);
                            //Application.Current.Shutdown();
                        }
                    }
                }
            }
        }

        //

        private void userName_TextChanged(object sender, TextChangedEventArgs e)
        {
            userName.Focus();
        }

        //

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Are you sure you want to Exit? ", "DotBrgy", MessageBoxButton.YesNo, MessageBoxImage.Information, MessageBoxResult.No);
            if (result == MessageBoxResult.Yes)

            // Do this

            {
                Application.Current.Shutdown();
            }
            else
            {
                return;
            }
        }

        private void startCamara_Checked(object sender, RoutedEventArgs e)
        {

        }
    }
}
