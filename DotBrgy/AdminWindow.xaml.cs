using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Word;
using WordWindow = Microsoft.Office.Interop.Word.Window;
using Window = System.Windows.Window;
using System.Xml;
using System.Xml.Linq;
using ContentControl = System.Windows.Controls.ContentControl;
using System.Data.SQLite;
using System.Data;
using CheckBox = System.Windows.Controls.CheckBox;
using System.Globalization;
using System.Reflection;
using System.Diagnostics;
using Path = System.IO.Path;
using System.IO.Packaging;
using Application = System.Windows.Application;
using System.Windows.Threading;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Mail;
using MailMessage = System.Net.Mail.MailMessage;
using DataTable = System.Data.DataTable;

namespace DotBrgy
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class AdminWindow : Window
    {

        public string count_item;
        //sqlite connection

        readonly string sqliteConnectionString = @"Data Source=C:\DotBrgy\Database\dotBrgyDatabase.sqlite;Version=3;";

        public AdminWindow()
        {
            InitializeComponent();
        }

        //
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //hide

            //button

            CivilStatus.Items.Add("Married");
            CivilStatus.Items.Add("Single");

            Sex.Items.Add("Male");
            Sex.Items.Add("Female");

            registeredVoter.Items.Add("Yes");
            registeredVoter.Items.Add("No");

            faucet.Items.Add("Yes");
            faucet.Items.Add("No");

            comfortRoom.Items.Add("Yes");
            comfortRoom.Items.Add("No");

            seniorID.Items.Add("Yes");
            seniorID.Items.Add("No");

            pwd.Items.Add("Yes");
            pwd.Items.Add("No");

            indigenousPeople.Items.Add("Yes");
            indigenousPeople.Items.Add("No");

            DisplayData();
            DisplayUserData();
        }

        //

        private void DisplayData()
        {
            try
            {
                if (!Directory.Exists(@"C:\DotBrgy"))
                {
                    return;
                }
                else
                {
                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        //da = new SQLiteDataAdapter("Select * From Student order by ID desc", con);
                        using (SQLiteDataAdapter sda = new SQLiteDataAdapter("Select * From dotBrgyData order by NOS desc", con))
                        {
                            using (DataSet dts = new DataSet())
                            {
                                con.Open();
                                sda.Fill(dts, "dotBrgyData");

                                dbData.ItemsSource = dts.Tables["dotBrgyData"].DefaultView;

                                //countTotal.Content = "Total: " + dbData.Items.Count.ToString();

                                countTotal.Content = dbData.Items.Count.ToString();

                                countMan.Content = count_item + dbData.Items.Cast<DataRowView>().Count(r => r.Row.ItemArray[11].ToString() == "Male");

                                countWoman.Content = count_item + dbData.Items.Cast<DataRowView>().Count(r => r.Row.ItemArray[11].ToString() == "Female");


                                //count_complete.Content = count_item + db_transaction.Items.Cast<DataRowView>().Count(r => r.Row.ItemArray[22].ToString() == "Completed");

                                //count_pending.Content = count_item + db_transaction.Items.Cast<DataRowView>().Count(r => r.Row.ItemArray[22].ToString().Trim() == "Pending");

                                //count_pending.Content = "Pending: " + count_item + db_transaction.Items.Cast<DataRowView>().Count(r => r.Row.ItemArray[22].ToString().Trim() == "Pending");

                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
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

        private void WaitCursor()
        {
            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait; // set the cursor to loading spinner  
        }

        private void NormalCursor()
        {
            Mouse.OverrideCursor = null; // set the cursor back to normal
        }

        //

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            Clear(this);
            DisplayData();
        }

        //

        private void Minimize_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void ButtonClose_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Are you sure you want to Logout? ", "DotBrgy", MessageBoxButton.YesNo, MessageBoxImage.Information, MessageBoxResult.No);
            if (result == MessageBoxResult.Yes)

            // Do this

            {
                LoginWindow open = new LoginWindow();
                open.Show();
                this.Hide();
                //Application.Current.Shutdown();
                //KillWord();
            }
            else
            {
                return;
            }
        }

        private void dbData_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                DataGrid gd = (DataGrid)sender;
                if (gd.SelectedItem is DataRowView row_selected)
                {
                    WaitCursor();
                    trackId.Text = row_selected["NOS"].ToString();
                    brgyID.Text = row_selected["brgyID"].ToString();
                    firstname.Text = row_selected["firstname"].ToString();
                    middlename.Text = row_selected["middlename"].ToString();
                    lastname.Text = row_selected["lastname"].ToString();
                    purok.Text = row_selected["purok"].ToString();
                    houseNumber.Text = row_selected["houseNumber"].ToString();
                    CivilStatus.Text = row_selected["CivilStatus"].ToString();
                    Sex.Text = row_selected["Sex"].ToString();
                    Birthdate.Text = row_selected["Birthdate"].ToString();
                    birthPlace.Text = row_selected["BirthPlace"].ToString();
                    educationAttainment.Text = row_selected["EducationalAttainment"].ToString();
                    philHealth.Text = row_selected["Philhealth"].ToString();
                    registeredVoter.Text = row_selected["RegisteredVoter"].ToString();
                    occupation.Text = row_selected["Occupation"].ToString();
                    familyPlanning.Text = row_selected["FamilyPlanning"].ToString();
                    faucet.Text = row_selected["Faucet"].ToString();
                    comfortRoom.Text = row_selected["ComfortRoom"].ToString();
                    seniorID.Text = row_selected["SeniorID"].ToString();
                    pwd.Text = row_selected["PWD"].ToString();
                    indigenousPeople.Text = row_selected["IndigenousPeople"].ToString();
                    membership.Text = row_selected["Membership"].ToString();
                    //Birthdate.Text = DateTime.Now.ToString();
                    //age.Text = row_selected["age"].ToString();
                    NormalCursor();

                    //DataView dv = dbData.ItemsSource as DataView;
                    //dv.RowFilter = "Convert(residentCode, 'System.String') like '%" + coded + "%'"; //where n is a column name of the DataTable

                    var years = DateTime.Now.Year - Birthdate.SelectedDate.Value.Year;

                    if (Birthdate.SelectedDate.Value.AddYears(years) > DateTime.Now) years--;
                    {
                        age.Text = years.ToString();
                    }

                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        con.Open();
                        using (SQLiteCommand cmd = new SQLiteCommand("Select residentCode from dotBrgyData where NOS like @nos", con))
                        {
                            cmd.Parameters.AddWithValue("@nos", trackId.Text);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                   string code = reader["residentCode"].ToString();
                                   DataView dv = dbData.ItemsSource as DataView;
                                   dv.RowFilter = "Convert(residentCode, 'System.String') like '%" + code + "%'"; //where n is a column name of the DataTable
                                   return;
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception)
            {
                throw;
            }
        }

        private void TrackId_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        //

        private void OpenMain_MouseLeftButtonUp_2(object sender, MouseButtonEventArgs e)
        {
            WaitCursor();
            MainWindow open = new MainWindow();
            open.Show();
            this.Hide();
            NormalCursor();
        }

        private void OpenTech_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            WaitCursor();
            openTechnical.IsOpen = true;
            NormalCursor();
        }

        private void Exit_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Are you sure you want to Exit? ", "DotBrgy", MessageBoxButton.YesNo, MessageBoxImage.Information, MessageBoxResult.No);
            if (result == MessageBoxResult.Yes)

            // Do this

            {
                Application.Current.Shutdown();
                //KillWord();
            }
            else
            {
                return;
            }
        }

        private void Search_TextChanged(object sender, TextChangedEventArgs e)
        {
            Regex rx = new Regex("^[a-zA-Z]+$");
            if (rx.IsMatch(search.Text))
            {
                DataView dv = dbData.ItemsSource as DataView;
                dv.RowFilter = string.Format("lastname LIKE '%{0}%' or purok LIKE '{0}%'", search.Text); //where n is a column name of the DataTable
            }
            else
            {
                return;
            }
        }

        private void DeleteAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dbData.Items.Count == 0)
                {
                    MessageBox.Show(this, "No data found!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                else
                {
                    MessageBoxResult delete_all = MessageBox.Show("Are you sure you want to delete all the data? This cannot be undone!", "Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
                    if (delete_all == MessageBoxResult.Yes)
                    {

                        using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                        {
                            using (SQLiteCommand deleteTransaction = con.CreateCommand())
                            {
                                using (SQLiteCommand cleanTransaction = con.CreateCommand())
                                {
                                    using (SQLiteCommand deleteStat = con.CreateCommand())
                                    {
                                        using (SQLiteCommand cleanStat = con.CreateCommand())
                                        {
                                            using (SQLiteCommand deleteHistory = con.CreateCommand())
                                            {
                                                using (SQLiteCommand cleanHistory = con.CreateCommand())
                                                {
                                                    con.Open();
                                                    //cmd.CommandType = CommandType.Text;
                                                    //cmd.CommandText = "select distinct address from tb_address_idType";
                                                    //cmd.CommandText = "DELETE FROM sqlite_sequence WHERE name = '%transactions%'";
                                                    //cmd.CommandText = "UPDATE sqlite_sequence SET seq = 10 WHERE name = 'transactions'";
                                                    //cmd.CommandText = "truncate table transactions";
                                                    //cmd.CommandText = "delete from [transactions]";

                                                    deleteTransaction.CommandText = "delete from dotBrgyData";
                                                    deleteTransaction.ExecuteNonQuery();

                                                    cleanTransaction.CommandText = "DELETE FROM sqlite_sequence WHERE name = 'dotBrgyData'";
                                                    cleanTransaction.CommandText = "UPDATE sqlite_sequence SET seq = 0 WHERE name = 'dotBrgyData'";
                                                    cleanTransaction.ExecuteNonQuery();


                                                    deleteStat.CommandText = "delete from statistic";
                                                    deleteStat.ExecuteNonQuery();

                                                    cleanStat.CommandText = "DELETE FROM sqlite_sequence WHERE name = 'statistic'";
                                                    cleanStat.CommandText = "UPDATE sqlite_sequence SET seq = 0 WHERE name = 'statistic'";
                                                    cleanStat.ExecuteNonQuery();

                                                    deleteHistory.CommandText = "delete from history";
                                                    deleteHistory.ExecuteNonQuery();

                                                    cleanHistory.CommandText = "DELETE FROM sqlite_sequence WHERE name = 'history'";
                                                    cleanHistory.CommandText = "UPDATE sqlite_sequence SET seq = 0 WHERE name = 'history'";
                                                    cleanHistory.ExecuteNonQuery();

                                                    DisplayData();
                                                    //Clean();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else if (delete_all == MessageBoxResult.No)
                    {
                        return;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }

        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            if (dbData.Items.Count == 0)
            {
                MessageBox.Show(this, "No data found!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else if (trackId.Text == string.Empty)
            {
                MessageBox.Show("Select valid item!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else
            {
                MessageBoxResult result = MessageBox.Show("Are you sure you want to delete? This cannot be undone!", "Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
                if (result == MessageBoxResult.Yes)

                // Do this

                {

                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        using (SQLiteCommand cmd = con.CreateCommand())
                        {
                            con.Open();
                            cmd.CommandType = CommandType.Text;
                            //cmd.CommandText = "alter table transactions AUTO_INCREMENT = 1";
                            //cmd.CommandText = "truncate table transactions";
                            //cmd.CommandText = "delete from [transactions]";
                            cmd.CommandText = "delete from dotBrgyData where NOS=@nos";
                            cmd.Parameters.AddWithValue("@nos", trackId.Text);
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Deleted!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            DisplayData();
                            Clear(this);
                        }
                    }
                }
            }
        }

        private void age_TextChanged(object sender, TextChangedEventArgs e)
        {
            /*
            var years = DateTime.Now.Year - Birthdate.SelectedDate.Value.Year;

            if (Birthdate.SelectedDate.Value.AddYears(years) > DateTime.Now) years--;
            {
                age.Text = years.ToString();
            }
            */
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            if (dbData.Items.Count == 0)
            {
                MessageBox.Show(this, "No data found!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else if (trackId.Text == string.Empty)
            {
                MessageBox.Show("Select valid item!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else
            {
                using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                {
                    using (SQLiteCommand cmd = con.CreateCommand())
                    {
                        WaitCursor();
                        con.Open();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "update dotBrgyData set img=@img,brgyID=@brgyID,firstname=@firstname,middlename=@middlename,lastname=@lastname,purok=@purok,Sex=@Sex,Birthdate=@Birthdate,BirthPlace=@BirthPlace,CivilStatus=@CivilStatus,EducationalAttainment=@EducationalAttainment," +
                            "Philhealth=@Philhealth,RegisteredVoter=@RegisteredVoter,Occupation=@Occupation,FamilyPlanning=@FamilyPlanning,Faucet=@Faucet,ComfortRoom=@ComfortRoom,SeniorID=@SeniorID,PWD=@PWD,IndigenousPeople=@IndigenousPeople,Membership=@Membership where NOS=" + trackId.Text;
                        //cmd.CommandText = "update transactions  set ownerFirstname=@firstname,ownerMiddlename=@middlename,ownerLastname=@lastname,CivilStatus=@CivilStatus,succeedingAction=@action where id=" + TrackId.Text;

                        cmd.Parameters.AddWithValue("@firstname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(firstname.Text));
                        cmd.Parameters.AddWithValue("@middlename", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(middlename.Text));
                        cmd.Parameters.AddWithValue("@lastname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(lastname.Text));

                        cmd.Parameters.AddWithValue("@CivilStatus", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(CivilStatus.Text));
                        cmd.Parameters.AddWithValue("@purok", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(purok.Text));
                        cmd.Parameters.AddWithValue("@Sex", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(Sex.Text));

                        cmd.Parameters.AddWithValue("Birthdate", Birthdate.Text);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Record has been successfully updated!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Question);
                        DisplayData();
                        //display_transaction();
                        Clear(this);
                        NormalCursor();
                    }
                }
            }
        }

        ///USER ACCOUNT

        private void DisplayUserData()
        {
            if (!Directory.Exists(@"C:\DotBrgy"))
            {
                return;
            }
            else
            {
                using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                {
                    //da = new SQLiteDataAdapter("Select * From Student order by ID desc", con);
                    using (SQLiteDataAdapter sda = new SQLiteDataAdapter("Select * From userAccount order by NOS desc", con))
                    {
                        using (DataSet dts = new DataSet())
                        {
                            con.Open();
                            sda.Fill(dts, "dotBrgyData");
                            db_userAccount.ItemsSource = dts.Tables["dotBrgyData"].DefaultView;
                        }
                    }
                }
            }
        }

        private void SaveUser_Click(object sender, RoutedEventArgs e)
        {
            Regex rx = new Regex("[^a-zA-Z0-9_.]+");
            //Regex rx_char = new Regex(@".{8,15}");
            if (rx.IsMatch(user_username.Text) || rx.IsMatch(user_password.Text))
            {
                MessageBox.Show(this, "Not accepting special character's!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            /*else if (!rx_char.IsMatch(tbx_username.Text) || !rx_char.IsMatch(tbx_password.Text))
            {
                MessageBox.Show(this, "Must contain at least 8 character!", "DotBrgy", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }*/
            else if (user_checkboxAdmin.IsChecked == false && user_checkboxStandard.IsChecked == false)
            {
                MessageBox.Show(this, "Please select atleast 1 user type!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }
            else if (user_checkboxAdmin.IsChecked == true)
            {
                using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                {
                    using (SQLiteCommand cmd = con.CreateCommand())
                    {
                        con.Open();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "insert into userAccount(userType, username, password, firstname, lastname, address) values(@userType,@username,@password,@firstname,@lastname,@address)";

                        //save image

                        //MemoryStream stream = new MemoryStream();
                        //through the instruction below, we save the
                        //image to byte in the object "stream".
                        //user_image.Image.Save(stream, System.Drawing.Imaging.ImageFormat.Jpeg);

                        //Below is the most important part, actually you are
                        //transferring the bytes of the array
                        //to the pic which is also of kind byte[]
                        //byte[] pic = stream.ToArray();

                        cmd.Parameters.AddWithValue("@userType", user_checkboxAdmin.Content);
                        cmd.Parameters.AddWithValue("@username", user_username.Text);
                        cmd.Parameters.AddWithValue("@password", user_password.Text);
                        cmd.Parameters.AddWithValue("@firstname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(user_firstname.Text));
                        cmd.Parameters.AddWithValue("@lastname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(user_lastname.Text));
                        cmd.Parameters.AddWithValue("@address", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(user_address.Text));


                        using (SQLiteCommand command = new SQLiteCommand("Select count (*) from userAccount where username = '" + user_username.Text + "'", con))
                        {
                            using (SQLiteCommand command1 = new SQLiteCommand("Select count (*) from userAccount where password = '" + user_password.Text + "'", con))
                            {
                                var result = command.ExecuteScalar();
                                int i = Convert.ToInt32(result);
                                var result1 = command1.ExecuteScalar();
                                int i1 = Convert.ToInt32(result1);
                                if (i != 0 || i1 != 0)
                                {
                                    MessageBox.Show("Username or Password already exist!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Information); ;
                                    Clear(this);
                                    return;
                                }
                                else
                                {
                                    cmd.ExecuteNonQuery();
                                    MessageBox.Show("User save in database!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Information);
                                    DisplayUserData();
                                    Clear(this);
                                }
                            }
                        }
                    }
                }
            }
            else if (user_checkboxStandard.IsChecked == true)
            {
                using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                {
                    using (SQLiteCommand cmd = con.CreateCommand())
                    {
                        con.Open();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "insert into userAccount(userType, username, password, firstname, lastname, address) values(@userType,@username,@password,@firstname,@lastname,@address)";

                        //save image

                        //MemoryStream stream = new MemoryStream();
                        //through the instruction below, we save the
                        //image to byte in the object "stream".
                        //user_image.Image.Save(stream, System.Drawing.Imaging.ImageFormat.Jpeg);

                        //Below is the most important part, actually you are
                        //transferring the bytes of the array
                        //to the pic which is also of kind byte[]
                        //byte[] pic = stream.ToArray();

                        cmd.Parameters.AddWithValue("@userType", user_checkboxStandard.Content);
                        cmd.Parameters.AddWithValue("@username", user_username.Text);
                        cmd.Parameters.AddWithValue("@password", user_password.Text);
                        cmd.Parameters.AddWithValue("@firstname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(user_firstname.Text));
                        cmd.Parameters.AddWithValue("@lastname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(user_lastname.Text));
                        cmd.Parameters.AddWithValue("@address", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(user_address.Text));

                        using (SQLiteCommand command = new SQLiteCommand("Select count (*) from userAccount where username = '" + user_username.Text + "'", con))
                        {
                            using (SQLiteCommand command1 = new SQLiteCommand("Select count (*) from userAccount where password = '" + user_password.Text + "'", con))
                            {
                                var result = command.ExecuteScalar();
                                int i = Convert.ToInt32(result);
                                var result1 = command.ExecuteScalar();
                                int i1 = Convert.ToInt32(result1);
                                if (i != 0 || i1 != 0)
                                {
                                    MessageBox.Show("Username or Password already exist!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Information); ;
                                    Clear(this);
                                    return;
                                }
                                else
                                {
                                    cmd.ExecuteNonQuery();
                                    MessageBox.Show("User save in database!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Information);
                                    DisplayUserData();
                                    Clear(this);
                                }
                            }
                        }
                    }
                }
            }
        }

        private void EditUser_Click(object sender, RoutedEventArgs e)
        {
            if (userId.Text == string.Empty)
            {
                MessageBox.Show("Select valid item!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                {
                    using (SQLiteCommand cmd = con.CreateCommand())
                    {
                        DataRowView drv = (DataRowView)db_userAccount.SelectedItem;
                        con.Open(); cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "update userAccount set firstname=@firstname,lastname=@lastname,username=@username,password=@password where NOS=" + userId.Text;
                        cmd.Parameters.AddWithValue("@firstname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(user_firstname.Text));
                        cmd.Parameters.AddWithValue("@lastname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(user_lastname.Text));
                        cmd.Parameters.AddWithValue("@username", user_username.Text);
                        cmd.Parameters.AddWithValue("@password", user_password.Text);

                        cmd.ExecuteNonQuery();
                        DisplayUserData();
                        MessageBox.Show("Record has been successfully updated!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Question);
                    }
                }
            }
        }

        private void UserRefresh_Click(object sender, RoutedEventArgs e)
        {
            WaitCursor();
            DisplayUserData();
            Clear(this);
            NormalCursor();
        }

        private void UserDelete_Click(object sender, RoutedEventArgs e)
        {
            if (userId.Text == string.Empty)
            {
                MessageBox.Show("Select valid item!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                MessageBoxResult result = MessageBox.Show("Are you sure you want to delete? This cannot be undone!", "Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
                if (result == MessageBoxResult.Yes)

                // Do this

                {

                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        using (SQLiteCommand cmd = con.CreateCommand())
                        {
                            con.Open();
                            cmd.CommandType = CommandType.Text;
                            //cmd.CommandText = "alter table transactions AUTO_INCREMENT = 1";
                            //cmd.CommandText = "truncate table transactions";
                            //cmd.CommandText = "delete from [transactions]";
                            cmd.CommandText = "delete from userAccount where NOS=@nos";
                            cmd.Parameters.AddWithValue("@nos", userId.Text);
                            cmd.ExecuteNonQuery();
                            DisplayUserData();
                            //MessageBox.Show("Deleted!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            Clear(this);
                        }
                    }
                }
            }
        }

        private void Db_userAccount_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {

                //MessageBox.Show(this, "Successfully Login!", "DotBrgy", MessageBoxButtons.OK, MessageBoxIcon.Information);   
                //username = textbox_username.Text;
                DataGrid gd = (DataGrid)sender;
                if (gd.SelectedItem is DataRowView row_selected)
                {
                    userId.Text = row_selected["NOS"].ToString();
                    user_firstname.Text = row_selected["firstname"].ToString();
                    user_lastname.Text = row_selected["lastname"].ToString();
                    user_username.Text = row_selected["username"].ToString();
                    user_password.Text = row_selected["password"].ToString();
                    user_address.Text = row_selected["address"].ToString();
                    //user_checkboxAdmin.IsChecked = true;

                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        using (SQLiteCommand cmd_username = con.CreateCommand())
                        {
                            //wait_cursor();
                            con.Open();
                            cmd_username.CommandType = CommandType.Text;
                            cmd_username.CommandText = "Select * from userAccount where username = '" + user_username.Text.Trim() + "'and password = '" + user_password.Text.Trim() + "'";
                            SQLiteDataReader sdr_admin;
                            sdr_admin = cmd_username.ExecuteReader();
                            //int count_admin = 0;
                            string userRole_admin = string.Empty;

                            while (sdr_admin.Read())
                            {
                                //count_admin = count_admin + 1;
                                userRole_admin = sdr_admin["userType"].ToString();

                            }


                            if (userRole_admin == "Administrator")
                            {
                                user_checkboxAdmin.IsChecked = true;
                            }
                            else if (userRole_admin == "Standard")
                            {
                                user_checkboxStandard.IsChecked = true;
                            }
                        }
                    }
                }
            }
            catch (System.Exception)
            {
                MessageBox.Show("Please call Technical Support!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SearchUser_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                Regex rx = new Regex("^[a-zA-Z]+$");
                if (rx.IsMatch(searchUser.Text))
                {
                    DataView dv = db_userAccount.ItemsSource as DataView;
                    //dv.RowFilter = string.Format("lastname LIKE '%{0}%'", userId.Text); //where n is a column name of the DataTable
                    dv.RowFilter = string.Format("lastname LIKE '%{0}%' or firstname LIKE '{0}%'", searchUser.Text);
                }
                else
                {
                    DataView dv = db_userAccount.ItemsSource as DataView;
                    dv.RowFilter = "Convert(NOS, 'System.String') like '%" + searchUser.Text + "%'"; //where n is a column name of the DataTable
                }
            }
            catch (System.Exception)
            {
                MessageBox.Show("No data found!");
            }
        }

        private void resCodeRadio_Checked(object sender, RoutedEventArgs e)
        {
            DataView dv = dbData.ItemsSource as DataView;
            dv.RowFilter = "Convert(residentCode, 'System.String') like '%" + search.Text + "%'"; //where n is a column name of the DataTable                                                                                                           
        }
        private void brgyIDradio_Checked(object sender, RoutedEventArgs e)
        {
            DataView dv = dbData.ItemsSource as DataView;
            dv.RowFilter = "Convert(brgyID, 'System.String') like '%" + search.Text + "%'"; //where n is a column name of the DataTable                                                                                                 
        }
        private void ageRadio_Checked(object sender, RoutedEventArgs e)
        {
            DataView dv = dbData.ItemsSource as DataView;
            dv.RowFilter = "Convert(age, 'System.String') like '%" + search.Text + "%'"; //where n is a column name of the DataTable     
        }
        private void houseNumberRadio_Checked(object sender, RoutedEventArgs e)
        {
            DataView dv = dbData.ItemsSource as DataView;
            dv.RowFilter = "Convert(houseNumber, 'System.String') like '%" + search.Text + "%'"; //where n is a column name of the DataTable
        }
    }
}