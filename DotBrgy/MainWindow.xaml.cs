using AForge.Video;
using AForge.Video.DirectShow;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using ZXing;
using ZXing.Common;
using Color = System.Drawing.Color;
using Path = System.IO.Path;

namespace DotBrgy
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //
        public DataSet ds;
        public string count_item;

        VideoCaptureDevice LocalWebCam;
        public FilterInfoCollection LocalWebCamsCollection;
        private BitmapImage latestFrame;


        //excel
        public Microsoft.Office.Interop.Excel.Application excel = null;
        public Microsoft.Office.Interop.Excel.Workbook wb = null;
        public Microsoft.Office.Interop.Excel.Worksheet ws = null;

        //docx
        public Microsoft.Office.Interop.Word.Application wordApp = null;
        public Microsoft.Office.Interop.Word.Document myWordDoc = null;

        //sqlite connection

        readonly string sqliteConnectionString = @"Data Source=C:\DotBrgy\Database\dotBrgyDatabase.sqlite;Version=3;";

        //automatic backup
        readonly DispatcherTimer timerBackup = new DispatcherTimer();

        private readonly BackgroundWorker worker = new BackgroundWorker();

        //timer shutdown
        //readonly DispatcherTimer timerShut = new DispatcherTimer();

        public DateTime StartDate;
        public DateTime EndDate;
        public TimeSpan Difference;

        public MainWindow()
        {
            InitializeComponent();
            //timerShut.Interval = TimeSpan.FromMinutes(15);
            //timerShut.Tick += TickShutdown;
            //timerShut.Start();

            timerBackup.Interval = TimeSpan.FromSeconds(10);
            timerBackup.Tick += TickBackup;
            timerBackup.Start();
            worker.DoWork += Worker_DoWork;

            Loaded += CameraWindow_Loaded;
            Unloaded += CameraWindow_Unloaded;
        }

        private void CameraWindow_Loaded(object sender, RoutedEventArgs e)
        {
            LocalWebCamsCollection = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            LocalWebCam = new VideoCaptureDevice(LocalWebCamsCollection[0].MonikerString);
            LocalWebCam.VideoResolution = LocalWebCam.VideoCapabilities[0];
            LocalWebCam.NewFrame += new NewFrameEventHandler(Cam_NewFrame);

            //LocalWebCam.Start();
        }

        private void CameraWindow_Unloaded(object sender, RoutedEventArgs e)
        {
            LocalWebCam.Stop();
        }



        void Cam_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            try
            {
                /**/
                System.Drawing.Image img = (System.Drawing.Bitmap)eventArgs.Frame.Clone();

                MemoryStream ms = new MemoryStream();
                img.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                ms.Seek(0, SeekOrigin.Begin);
                BitmapImage bi = new BitmapImage();
                bi.BeginInit();
                bi.StreamSource = ms;
                bi.EndInit();
                bi.Freeze();
                this.latestFrame = bi;

                Dispatcher.BeginInvoke(new ThreadStart(delegate
                {
                    uploadImage.Source = bi;
                    uploadImage.Source = new CroppedBitmap(bi, new Int32Rect(400, 0, 400, 400));
                }));

            }
            catch (Exception)
            {
                throw;
            }
        }

        /*
        private void TickShutdown(object sender, EventArgs e)
        {
            Application.Current.Shutdown();
        }
        */

        private void TickBackup(object sender, EventArgs e)
        {
            worker.RunWorkerAsync();
        }

        //automatic backup

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            // run all background tasks here
            //time.Text = DateTime.Now.ToString("HH:mm:ss");
            //automatic backup
            try
            {
                if (dbData.Items.Count == 0)
                {
                    return;
                }
                else
                {
                    using (var source = new SQLiteConnection(@"Data Source=C:\DotBrgy\Database\dotBrgyDatabase.sqlite;Version=3"))
                    using (var destination = new SQLiteConnection(@"Data Source=C:\DotBrgy\Backup\Database\dotBrgyDatabase.sqlite;Version=3"))
                    //using (var hiddenDestination = new SQLiteConnection(@"Data Source=C:\Warning\DotBrgy\Database\Backup\TransaksyonTracerDatabase.sqlite"))
                    {
                        {
                            source.Open();
                            //C
                            destination.Open();
                            source.BackupDatabase(destination, "main", "main", -1, null, 0);

                            //
                            //hiddenDestination.Open();
                            //source.BackupDatabase(destination, "main", "main", -1, null, 0);
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        //

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //hide

            //button
            //business.Visibility = Visibility.Collapsed;
            //generate.IsEnabled = false;

            CivilStatus.Items.Add("Married");
            CivilStatus.Items.Add("Single");
            CivilStatus.Items.Add("Widowed");
            CivilStatus.Items.Add("Separated");
            CivilStatus.Items.Add("Divorced");

            transType.Items.Add("Business");
            transType.Items.Add("Clearance");
            transType.Items.Add("Indigency");
            transType.Items.Add("Residency");

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
            DisplayPurok();
            DisplayCode();
            DisplayQrCode();
            //TrialAlgorithm();
            DisplayHistory();
        }

        //

        private void DisplayAll()
        {
            DisplayData();
            DisplayPurok();
            DisplayCode();
            DisplayQrCode();
            DisplayHistory();

            //
        }

        private void TrialAlgorithm()
        {
            if (trialLeft.Text == "0")
            {
                return;
            }
            else
            {
                if (!Directory.Exists(@"C:\DotBrgy"))
                {
                    return;
                }
                else
                {
                    /*
                    using (SQLiteDataAdapter sda = new SQLiteDataAdapter("Select dateBegin from trialBegin", sqliteConnectionString))
                    {
                        DataTable dt_code = new DataTable();
                        sda.Fill(dt_code);
                        //id_number.Text = "ID Number: " + dt.Rows[0][0].ToString();

                        dateBegin.Text = dt_code.Rows[0]["dateBegin"].ToString();
                        //return;
                        //dateEnd.Text = dt_code.Rows[1]["dateEnd"].ToString();
                    } */

                    /*
                    using (SQLiteDataAdapter sdaEnd = new SQLiteDataAdapter("Select dateEnd from trialEnd", sqliteConnectionString))
                    {
                        DataTable dtEnd = new DataTable();
                        sdaEnd.Fill(dtEnd);
                        //id_number.Text = "ID Number: " + dt.Rows[0][0].ToString();

                        dateEnd.Text = dtEnd.Rows[0]["dateEnd"].ToString();
                        //return;
                        //dateEnd.Text = dt_code.Rows[1]["dateEnd"].ToString();
                    }*/

                    using (SQLiteDataAdapter sdaEnd = new SQLiteDataAdapter("Select dateEnd from trialEnd", sqliteConnectionString))
                    {
                        using (SQLiteDataAdapter sdaBegin = new SQLiteDataAdapter("Select dateBegin from trialBegin", sqliteConnectionString))
                        {
                            //start
                            DataTable dtBegin = new DataTable();
                            sdaBegin.Fill(dtBegin);

                            //end
                            DataTable dtEnd = new DataTable();
                            sdaEnd.Fill(dtEnd);
                            //id_number.Text = "ID Number: " + dt.Rows[0][0].ToString();

                            StartDate = Convert.ToDateTime(dtBegin.Rows[0]["dateBegin"].ToString());
                            EndDate = Convert.ToDateTime(dtEnd.Rows[0]["dateEnd"].ToString());
                            //EndDate = Convert.ToDateTime(dateEnd.Text.ToString());

                            Difference = EndDate.Subtract(StartDate);

                            int total = Difference.Days;

                            trialLeft.Text = total.ToString();

                            if (total >= 50)
                            {
                                // disableTop.IsEnabled = false;
                                disableButtons.IsEnabled = false;
                                //this.IsEnabled = false;
                            }
                            else
                            {
                                if (!Directory.Exists(@"C:\DotBrgy"))
                                {
                                    return;
                                }
                                else
                                {
                                    using (SQLiteConnection conStat = new SQLiteConnection(sqliteConnectionString))
                                    {
                                        using (SQLiteCommand cmdStat = conStat.CreateCommand())
                                        {
                                            conStat.Open();
                                            cmdStat.CommandType = CommandType.Text;
                                            cmdStat.CommandText = "update trialEnd set dateEnd=@d";
                                            cmdStat.Parameters.AddWithValue("@d", DateTime.Now.ToString());
                                            cmdStat.ExecuteNonQuery();
                                            //MessageBox.Show("OK hakdog!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Information);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        //

        private void DisplayData()
        {
            if (!Directory.Exists(@"C:\DotBrgy"))
            {
                return;
            }
            else
            {
                using (SQLiteConnection conn = new SQLiteConnection(sqliteConnectionString))
                {
                    conn.Open();

                    using (SQLiteDataAdapter adapter = new SQLiteDataAdapter("SELECT * FROM dotBrgyData", conn))
                    {
                        ds = new DataSet();
                        adapter.Fill(ds);
                        DataTable dt = ds.Tables[0];

                        //cbImages.Items.Clear();

                        //foreach (DataRow dr in dt.Rows)
                        //cbImages.Items.Add(dr["id"].ToString());

                        //cbImages.SelectedIndex = 0;
                    }
                }

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

                            DateTime date;

                            foreach (DataRowView dr in dbData.Items)
                            {
                                date = Convert.ToDateTime(dr["Birthdate"].ToString());
                                int age1 = DateTime.Now.Year - date.Year;
                                dr["age"] = age1.ToString();
                            }


                            //count senior
                            int[] array = dbData.Items
                                 .Cast<DataRowView>()
                                 .Select(row => int.Parse(row.Row["age"].ToString()))
                                 .ToArray();

                            //int[] ar = { 3, 1, 2, 3, 4, 5, 6, 8};
                            //int max = ar[10];
                            int occurrenceCount = 1;

                            if (dbData.Items.Count == 0)
                            {
                                countElder.Content = 0;
                            }
                            else
                            {
                                for (int i = 0; i < array.Count(); i++)
                                {
                                    if (array[i] >= 60)
                                    {
                                        countElder.Content = occurrenceCount++;
                                    }
                                    else if (occurrenceCount == 1)
                                    {
                                        countElder.Content = 0;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        //

        private void DisplayCode()
        {
            if (dbData.Items.Count == 0)
            {
                code.Text = string.Empty;
            }
            else
            {
                try
                {    /**/
                    using (SQLiteDataAdapter sda = new SQLiteDataAdapter("Select NOS,residentCode from dotBrgyData ORDER BY NOS DESC", sqliteConnectionString))
                    {
                        DataTable dt_code = new DataTable();
                        sda.Fill(dt_code);
                        //id_number.Text = "ID Number: " + dt.Rows[0][0].ToString();
                        code.Text = dt_code.Rows[0]["residentCode"].ToString();
                    }
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }

        private void DisplayQrCode()
        {
            if (dbData.Items.Count == 0)
            {
                return;
            }
            else
            {
                //display qrcode

                //var qrcode = new QRCodeWriter();
                //var qrValue = "your magic here";

                var barcodeWriter = new BarcodeWriter
                {
                    Format = BarcodeFormat.QR_CODE,
                    Options = new EncodingOptions
                    {
                        Height = 300,
                        Width = 300,
                        Margin = 0,
                        PureBarcode = false
                    }
                };

                //string imageText = "DotBrgy";
                //string imageText = "";
                //Rectangle rectf = new Rectangle(85, 250, 0, 0);

                using (var bitmap = barcodeWriter.Write(code.Text))
                using (var stream = new MemoryStream())
                {
                    /*
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        using (Font arialFont = new Font("Century Gothic", 20))
                        {
                            using (StringFormat sf = new StringFormat())
                            {
                                //graphics
                                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                                graphics.DrawString(imageText, arialFont, Brushes.Black, rectf, sf);
                                  */
                    //bitmap
                    bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                    BitmapImage bi = new BitmapImage();
                    bi.BeginInit();
                    stream.Seek(0, SeekOrigin.Begin);
                    bi.StreamSource = stream;
                    bi.CacheOption = BitmapCacheOption.OnLoad;
                    bi.EndInit();
                    qrImage.Source = bi;
                }
            }
        }

        //

        private void DisplayPurok()
        {
            if (!Directory.Exists(@"C:\DotBrgy"))
            {
                return;
            }
            else
            {
                try
                {
                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        con.Open();

                        //address

                        DataTable dt = new DataTable();
                        using (SQLiteDataAdapter sda = new SQLiteDataAdapter("select distinct purok from dotBrgyData", con))
                        {
                            sda.Fill(dt);

                            purok.DisplayMemberPath = "purok";
                            purok.ItemsSource = dt.DefaultView;
                        }
                    }
                }
                catch (System.Exception)
                {
                    throw;
                }
            }
        }

        //display history

        private void DisplayHistory()
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
                    using (SQLiteDataAdapter sda = new SQLiteDataAdapter("Select * From history order by NOS desc", con))
                    {
                        using (DataSet dts = new DataSet())
                        {
                            con.Open();
                            sda.Fill(dts, "history");
                            dbHistory.ItemsSource = dts.Tables["history"].DefaultView;

                            //countTotal.Content = "Total: " + dbData.Items.Count.ToString();

                            //countTotal.Content = dbData.Items.Count.ToString();

                            //countMan.Content = count_item + dbData.Items.Cast<DataRowView>().Count(r => r.Row.ItemArray[8].ToString() == "Male");

                            //countWoman.Content = count_item + dbData.Items.Cast<DataRowView>().Count(r => r.Row.ItemArray[8].ToString() == "Female");
                        }
                    }
                }
            }
        }

        //Clear Data

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
                if (obj is DatePicker datepick)
                    datepick.SelectedDate = DateTime.Now;

                Clear(VisualTreeHelper.GetChild(obj, i));
            }
        }


        //mouse cursor

        private void WaitCursor()
        {
            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait; // set the cursor to loading spinner  
        }

        //

        private void NormalCursor()
        {
            Mouse.OverrideCursor = null; // set the cursor back to normal
        }

        //Find and Replace Method
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            //object read_only = false;
            //object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

        //Creeate the Doc Method
        private void CreateWordDocument(object filename, object SaveAs)
        {
            try
            {

                wordApp = new Microsoft.Office.Interop.Word.Application();
                object missing = Missing.Value;
                Microsoft.Office.Interop.Word.Document myWordDoc = null;

                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;



                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                           ref missing, ref missing, ref missing,
                                           ref missing, ref missing, ref missing,
                                           ref missing, ref missing, ref missing,
                                           ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();

                    //find and replace
                    this.FindAndReplace(wordApp, "<firstname>", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(firstname.Text));
                    this.FindAndReplace(wordApp, "<middlename>", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(middlename.Text));
                    this.FindAndReplace(wordApp, "<lastname>", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(lastname.Text));
                    this.FindAndReplace(wordApp, "<age>", age.Text);
                    this.FindAndReplace(wordApp, "<CivilStatus>", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(CivilStatus.Text));
                    this.FindAndReplace(wordApp, "<purok>", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(purok.Text));
                    this.FindAndReplace(wordApp, "<date>", DateTime.Now.ToString("MMMM dd, yyyy"));
                    this.FindAndReplace(wordApp, "<day>", DateTime.Now.ToString("dd"));
                    this.FindAndReplace(wordApp, "<month>", DateTime.Now.ToString("MMMM"));
                    this.FindAndReplace(wordApp, "<year>", DateTime.Now.ToString("yyyy"));
                    this.FindAndReplace(wordApp, "<or>", DateTime.Now.ToString("yyyyMMdd-HHmmss-fff"));
                    this.FindAndReplace(wordApp, "<business>", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(businessName.Text));
                    //this.FindAndReplace(wordApp, "<date>", DateTime.Now.ToShortDateString());
                }
                else
                {
                    MessageBox.Show("File not Found!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                //Save as
                myWordDoc.SaveAs(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                //insert image into document
                //myWordDoc.Bookmarks["add"].Range.InsertParagraph.myWordDoc

                //myWordDoc.Bookmarks["add"].Range.InsertParagraph.myWordDoc;

                //var shape = myWordDoc.InlineShapes.AddPicture(@"C:\DotBrgy\Images\eBrgyQR.png", false, true);

                //shape.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //shape.Width = 100;
                //shape.Height = 100;

                /*
                var outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
                var logoimage = Path.Combine(outPutDirectory, @"C:\DotBrgy\Images\eBrgyQR.png");

                var keyword = "LOGO";
                var sel = wordApp.Selection;
                sel.Find.Text = string.Format("[{0}]", keyword);
                wordApp.Selection.Find.Execute(keyword);

                Microsoft.Office.Interop.Word.Range range = wordApp.Selection.Range;
                if (range.Text.Contains(keyword))
                {
                    //gets desired range here it gets last character to make superscript in range 
                    Microsoft.Office.Interop.Word.Range temprange = myWordDoc.Range(range.End - 4, range.End);//keyword is of 4 charecter range.End - 4
                    temprange.Select();
                    Microsoft.Office.Interop.Word.Selection currentSelection = wordApp.Selection;
                    //currentSelection.Font.Superscript = 1;

                    sel.Find.Execute(Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne);
                    sel.Range.Select();
                    var imagePath1 = Path.GetFullPath(string.Format(logoimage, keyword));
                    var shape = sel.InlineShapes.AddPicture(FileName: imagePath1, LinkToFile: false, SaveWithDocument: true);
                    shape.Width = 100;
                    shape.Height = 100;
                }
                */

                myWordDoc.Close();
                wordApp.Quit();
                MessageBoxResult result = MessageBox.Show("File successfully saved in your disk!, Click yes to Open.", "DotBrgy", MessageBoxButton.YesNo, MessageBoxImage.Asterisk);
                if (result == MessageBoxResult.Yes)
                {
                    /*
                    System.Diagnostics.Process myProcess = new System.Diagnostics.Process();
                    myProcess.StartInfo.FileName = @"C:\Backup\print.docx";
                    myProcess.StartInfo.CreateNoWindow = true;
                    myProcess.Start();
                    */

                    System.Diagnostics.Process.Start(@"C:\DotBrgy\Print\print.docx");
                }
                else
                {
                    //KillWord();
                    //MessageBox.Show("File successfully saved in your disk!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    return;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error: " + ex.ToString(), "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Error);
                MessageBox.Show("Word cannot save this file because it is already open elsewhere.", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Error);
                KillWord();
            }
        }


        private void SaveQRCODEImage()
        {
            //path
            String filePath = @"C:\DotBrgy\Images\eBrgyQR.png";

            //save image
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create((BitmapSource)qrImage.Source));
            using (FileStream stream = new FileStream(filePath, FileMode.Create)) encoder.Save(stream);
        }

        //

        private void SaveResidentImage()
        {
            //path
            String filePath = @"C:\DotBrgy\Images\residentImage.png";

            //save image
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create((BitmapSource)uploadImage.Source));
            using (FileStream stream = new FileStream(filePath, FileMode.Create)) encoder.Save(stream);
        }

        private void SaveData_Click(object sender, RoutedEventArgs e)
        {
            /*
            int value;
            if (Int32.TryParse(age.Text, out value))
            {                                                       
                if (value < 18)
                {
                    MessageBox.Show(this, "Invalid Birth Year!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                else
                {  */
            if (Birthdate.Text == string.Empty)
            {
                MessageBox.Show(this, "Please select valid Birthdate!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else
            {
                if (uploadImage.Source == null)
                {
                    MessageBox.Show("Please insert new Image!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {

                        using (SQLiteCommand cmd = con.CreateCommand())
                        {
                            //string var;
                            //var = "Track number " + generateId.Text;
                            //int monthParse = month.SelectedIndex + 1;

                            WaitCursor();

                            //insertImageData();
                            SaveResidentImage();

                            string fileN = @"C:\DotBrgy\Images\residentImage.png";
                            //Initialize a file stream to read the image file
                            FileStream fs = new FileStream(fileN, FileMode.Open, FileAccess.Read);

                            //Initialize a byte array with size of stream
                            byte[] imgByteArr = new byte[fs.Length];

                            //Read data from the file stream and put into the byte array
                            fs.Read(imgByteArr, 0, Convert.ToInt32(fs.Length));

                            //Close a file stream
                            fs.Close();

                            con.Open();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "insert into dotBrgyData(img,residentCode,houseNumber,brgyID,dateRegister,firstname,middlename,lastname,purok,Sex,Birthdate,BirtPlace,CivilStatus,EducationalAttainment,Philhealth,RegisteredVoter,Occupation,FamilyPlanning,Faucet,ComfortRoom,SeniorID,PWD,IndigenousPeople,Membership)" +
                            " values(@img,@residentCode,@houseNumber,@brgyID,@dateRegister,@firstname,@middlename,@lastname,@purok,@Sex,@Birthdate,@BirthPlace,@CivilStatus,@EducationalAttainment,@Philhealth,@RegisteredVoter,@Occupation,@FamilyPlanning,@Faucet,@ComfortRoom,@SeniorID,@PWD,@IndigenousPeople,@Membership)";
                            /*
                             string query = "(Select count(*) from transactions where id = '" + id_number.Text.Trim() + "')";

                             using (SqlCommand cmda = new SqlCommand(query, con))
                             {

                                 //string query_address = "(Select count(*) from transactions where address = '" + combo_address.Text.Trim() + "')";


                                 //SqlCommand cmd_address = new SqlCommand(query_address, con_transaction);
                                 int count = (int)cmda.ExecuteScalar();
                                 if (count > 0)
                                 {
                                     MessageBox.Show(this, "Transaction number already exist, Please try again!", "DotBrgy", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                     return;
                                 }
                             */

                            //cmd.Parameters.AddWithValue("@transaction_code", s = dt.ToString("yyyy-") + uid.ToString());
                            //cmd.Parameters.AddWithValue("@transaction_code", DateTime.Now.ToString("yyyyMMddHHmmssfffffff"));





                            //cmd.Parameters.AddWithValue("@date_time", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                            //cmd.Parameters.AddWithValue("@date_time", DateTime.Now.ToString("MM/d/yyyy hh:mm:ss tt"));

                            //cmd.Parameters.AddWithValue("@residentCode", DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                            cmd.Parameters.AddWithValue("@residentCode", DateTime.Now.ToString("yyyyMMdd-HHmmss-fff"));
                            cmd.Parameters.AddWithValue("@houseNumber", houseNumber.Text);
                            cmd.Parameters.AddWithValue("@brgyID", brgyID.Text);
                            //cmd.Parameters.AddWithValue("@residentCode", DateTime.Now.Ticks.ToString());
                            cmd.Parameters.AddWithValue("@dateRegister", DateTime.Now.ToString());
                            cmd.Parameters.AddWithValue("@firstname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(firstname.Text));
                            cmd.Parameters.AddWithValue("@middlename", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(middlename.Text));
                            cmd.Parameters.AddWithValue("@lastname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(lastname.Text));

                            cmd.Parameters.AddWithValue("@purok", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(purok.Text));
                            cmd.Parameters.AddWithValue("@Sex", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(Sex.Text));
                            cmd.Parameters.AddWithValue("@Birthdate", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(Birthdate.Text));
                            cmd.Parameters.AddWithValue("@CivilStatus", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(CivilStatus.Text));

                            cmd.Parameters.AddWithValue("@EducationalAttainment", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(educationAttainment.Text));
                            cmd.Parameters.AddWithValue("@Philhealth", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(philHealth.Text));
                            cmd.Parameters.AddWithValue("@RegisteredVoter", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(registeredVoter.Text));
                            cmd.Parameters.AddWithValue("@Occupation", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(occupation.Text));
                            cmd.Parameters.AddWithValue("@FamilyPlanning", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(familyPlanning.Text));
                            cmd.Parameters.AddWithValue("@Faucet", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(faucet.Text));
                            cmd.Parameters.AddWithValue("@ComfortRoom", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(comfortRoom.Text));
                            cmd.Parameters.AddWithValue("@SeniorID", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(seniorID.Text));
                            cmd.Parameters.AddWithValue("@PWD", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(pwd.Text));
                            cmd.Parameters.AddWithValue("@IndigenousPeople", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(indigenousPeople.Text));
                            cmd.Parameters.AddWithValue("@Membership", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(membership.Text));
                            cmd.Parameters.AddWithValue("img", imgByteArr);

                            MessageBox.Show("Successfully saved into Database!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Asterisk);

                            cmd.ExecuteNonQuery();
                            NormalCursor();
                            Prediction();
                            DisplayAll();
                            startCamera.Visibility = Visibility.Visible;
                            LocalWebCam.Stop();
                            stopCamera.Visibility = Visibility.Collapsed;
                            cameras.Visibility = Visibility.Collapsed;
                            uploadImage.Source = null;
                            captureCamera.Visibility = Visibility.Collapsed;
                        }
                    }
                }
            }
        }

        private void Prediction()
        {
            int statNo = 0;

            int countResident = Convert.ToInt32(count_item + dbData.Items.Cast<DataRowView>().Count(r => r.Row.ItemArray[6].ToString() == purok.Text));

            //equals.Text = (Convert.ToInt32(nos.Text) + Convert.ToInt32(add.Text)).ToString();
            /*DataView dv = db_transaction.ItemsSource as DataView;
            dv.RowFilter = "Convert(ID, 'System.String') like '%" + searchId.Text + "%'"; //where n is a column name of the DataTable
               */

            // string sOne = countResident;
            //string[] sTwo = sOne.Split('+');
            //int a = 1;
            //for (int i = 0; i <= sTwo.Length - 1; i++)
            //{
            //    a = a + Convert.ToInt32(sTwo[i]);
            //}

            int equal = 1 + countResident;

            //equal.Text = a.ToString();
            using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
            {
                con.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("Select NOS from statistic where purok like @p", con))
                {
                    cmd.Parameters.AddWithValue("@p", purok.Text);
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            statNo = Convert.ToInt32(reader["NOS"].ToString());
                        }
                    }
                }
            }

            using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
            {
                /*
                using (SQLiteCommand cmd = con.CreateCommand())
                {
                    con.Open();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "insert into statistic(purok, residents) values(@purok,@residents)";
                    cmd.Parameters.AddWithValue("@purok", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(countPurok.Text));
                    cmd.Parameters.AddWithValue("@residents", equal.Text);
                      */
                using (SQLiteCommand command = new SQLiteCommand("Select count (*) from statistic where NOS = '" + statNo + "'", con))
                {
                    con.Open();
                    var result = command.ExecuteScalar();
                    int i = Convert.ToInt32(result);
                    if (i != 0)
                    {

                        using (SQLiteConnection conEdit = new SQLiteConnection(sqliteConnectionString))
                        {
                            using (SQLiteCommand cmdEdit = conEdit.CreateCommand())
                            {
                                WaitCursor();
                                conEdit.Open();
                                cmdEdit.CommandType = CommandType.Text;
                                cmdEdit.CommandText = "update statistic set residents=@r where NOS=" + statNo;
                                //cmd.CommandText = "update transactions  set ownerFirstname=@firstname,ownerMiddlename=@middlename,ownerLastname=@lastname,CivilStatus=@CivilStatus,succeedingAction=@action where id=" + TrackId.Text;

                                cmdEdit.Parameters.AddWithValue("@r", equal);

                                cmdEdit.ExecuteNonQuery();

                                MessageBox.Show("Record has been successfully updated!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Question);
                                //DisplayData();
                                //display_transaction();
                                DisplayAll();
                                //Clear(this);
                                NormalCursor();
                            }
                        }
                    }
                    else
                    {
                        return;
                    }
                }
            }
        }

        //

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            Clear(this);
            DisplayData();
            DisplayCode();
            /*
            if (uploadImage.Source != null)
            {
                LocalWebCam.Stop();
                uploadImage.Source = null;
                cameras.Visibility = Visibility.Collapsed;
                captureCamera.Visibility = Visibility.Collapsed;
                stopCamera.Visibility = Visibility.Collapsed;
                startCamera.Visibility = Visibility.Visible;
            }
            */
        }

        //

        private void KillWord()
        {
            try
            {
                wordApp.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
                Marshal.FinalReleaseComObject(wordApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.ToString(), "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            /*
            Process[] process = Process.GetProcesses();
            foreach (Process pr in process)
            {
                if (!(pr.ProcessName.IndexOf("WORD") < 0))
                {
                    pr.Kill();
                }
            }
            */
        }

        private void Minimize_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
            //this.WindowState = WindowState.Maximized;
        }

        private void Close_Click(object sender, RoutedEventArgs e)
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

        private void DbData_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                DataGrid gd = (DataGrid)sender;
                if (gd.SelectedItem is DataRowView row_selected)
                {
                    WaitCursor();

                    //code.Text = DateTime.Now.ToString("yyyyMMdd-HHmmss-fff");
                    trackId.Text = row_selected["NOS"].ToString();
                    code.Text = row_selected["residentCode"].ToString();
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
                    DisplayQrCode();
                    NormalCursor();

                    DataView dv = dbData.ItemsSource as DataView;
                    dv.RowFilter = "Convert(residentCode, 'System.String') like '%" + code.Text + "%'"; //where n is a column name of the DataTable

                    var years = DateTime.Now.Year - Birthdate.SelectedDate.Value.Year;

                    if (Birthdate.SelectedDate.Value.AddYears(years) > DateTime.Now) years--;
                    {
                        age.Text = years.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please call technical Support! " + "\n" + ex.ToString(), "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Error);
                NormalCursor();
            }
        }

        private void TrackId_TextChanged(object sender, TextChangedEventArgs e)
        {

            DataTable dataTable = ds.Tables[0];

            foreach (DataRow row in dataTable.Rows)
            {
                if (row[0].ToString() == trackId.Text)
                {
                    //Store binary data read from the database in a byte array
                    byte[] blob = (byte[])row[1];
                    MemoryStream stream = new MemoryStream();
                    stream.Write(blob, 0, blob.Length);
                    stream.Position = 0;

                    System.Drawing.Image img = System.Drawing.Image.FromStream(stream);
                    BitmapImage bi = new BitmapImage();
                    bi.BeginInit();

                    MemoryStream ms = new MemoryStream();
                    img.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                    ms.Seek(0, SeekOrigin.Begin);
                    bi.StreamSource = ms;
                    bi.EndInit();
                    uploadImage.Source = bi;
                }
            }

            if (trackId.Text == string.Empty)
            {
                printDocs.Visibility = Visibility.Collapsed;
                printID.Visibility = Visibility.Collapsed;
                printAll.Visibility = Visibility.Visible;
                saveData.IsEnabled = true;
                transType.Visibility = Visibility.Collapsed;
                cameras.Visibility = Visibility.Collapsed;
                uploadImage.Source = null;
                //age.Visibility = Visibility.Collapsed;
                //printDocument.Visibility = Visibility.Collapsed;
                //logo.Visibility = Visibility.Visible;
                //age.Visibility = Visibility.Collapsed;
            }
            else
            {
                printDocs.Visibility = Visibility.Visible;
                printID.Visibility = Visibility.Visible;
                printAll.Visibility = Visibility.Collapsed;
                saveData.IsEnabled = false;
                transType.Visibility = Visibility.Visible;
                cameras.Visibility = Visibility.Visible;
                //age.Visibility = Visibility.Visible;
                //printDocument.Visibility = Visibility.Visible;
                //age.Visibility = Visibility.Visible;
                //logo.Visibility = Visibility.Collapsed;

                //DataView dv = dbData.ItemsSource as DataView;
                //dv.RowFilter = "Convert(NOS, 'System.String') like '%" + trackId.Text + "%'"; //where n is a column name of the DataTable

                //DataView dv = dbData.ItemsSource as DataView;
                //dv.RowFilter = "Convert(residentCode, 'System.String') like '%" + code.Text + "%'"; //where n is a column name of the DataTable
            }
            /**/
        }

        private void OpenAdmin_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {

            MessageBoxResult result = MessageBox.Show("You need to login first to access administrator, Continue?", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Information, MessageBoxResult.No);
            if (result == MessageBoxResult.Yes)
            {
                this.Hide();
                LoginWindow open = new LoginWindow();
                open.Show();
            }
            else
            {
                return;
            }
            /*
            AdminWindow open = new AdminWindow();
            open.Show();
            this.Hide();
            */
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
            /*
            if(search.Text.Length == 2)
            {
                DataView dv = dbData.ItemsSource as DataView;
                dv.RowFilter = "Convert(age, 'System.String') like '%" + search.Text + "%'"; //where n is a column name of the DataTable
            }
            else if(search.Text.Length > 3)
            {
                DataView dv = dbData.ItemsSource as DataView;
                dv.RowFilter = "Convert(residentCode, 'System.String') like '%" + search.Text + "%'"; //where n is a column name of the DataTable  
            }     
            */

            //documents.Visibility = Visibility.Collapsed;
            /*
            Regex rx = new Regex("^[a-zA-Z]+$");
            if (rx.IsMatch(search.Text))
            {
                //countPurok.Text = count_item + dbData.Items.Cast<DataRowView>().Count(r => r.Row.ItemArray[6].ToString() == search.Text);

                //dv.RowFilter = string.Format("lastname LIKE '%{0}%'", searchId.Text); //where n is a column name of the DataTable
                DataView dv = dbData.ItemsSource as DataView;
                //dv.RowFilter = string.Format("lastname LIKE '%{0}%'", search.Text); //where n is a column name of the DataTable
                dv.RowFilter = string.Format("lastname LIKE '%{0}%' or purok LIKE '{0}%'", search.Text); //where n is a column name of the DataTable
            }
            else
            {
                DataView dv = dbData.ItemsSource as DataView;
                dv.RowFilter = "Convert(residentCode + brgyID + houseNumber + age, 'System.String') like '%" + search.Text + "%'"; //where n is a column name of the DataTable                                                                                                            //dv.RowFilter = string.Format("residentCode LIKE '%{0}%'", int.Parse(search.Text)); //where n is a column name of the DataTable
            }    
            if(search.Text == string.Empty)
            {
                MessageBox.Show("Please enter some data!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                search.Focus();
                return;
            }
            */

            Regex rx = new Regex("^[a-zA-Z]+$");
            if (rx.IsMatch(search.Text))
            {
                DataView dv = dbData.ItemsSource as DataView;
                dv.RowFilter = string.Format("lastname LIKE '%{0}%' or purok LIKE '{0}%'", search.Text); //where n is a column name of the DataTable
            }
            /*
            else
            {
                if (search.Text == string.Empty)
                {
                    search.Width = 660;
                    numericContainer.Visibility = Visibility.Hidden;
                }
                else
                {
                    search.Width = 300;
                    numericContainer.Visibility = Visibility.Visible;
                }
            }
            */
        }

        private void SelectedBirthdate_TextChanged(object sender, TextChangedEventArgs e)
        {
            //int year = DateTime.Now.Year - Birthdate.SelectedDate.Value.Year;
            //int month = DateTime.Now.Month - Birthdate.SelectedDate.Value.Month;

            //Birthdate.Text = string.Empty;

            /*
           var years = DateTime.Now.Year - Birthdate.SelectedDate.Value.Year;

           if (Birthdate.SelectedDate.Value.AddYears(years) > DateTime.Now) years--;
           {
               age.Text = years.ToString();
           } */
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            if (trackId.Text == string.Empty)
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
                            DisplayData();
                            con.Open();
                            cmd.CommandType = CommandType.Text;
                            //cmd.CommandText = "alter table transactions AUTO_INCREMENT = 1";
                            //cmd.CommandText = "truncate table transactions";
                            //cmd.CommandText = "delete from [transactions]";
                            cmd.CommandText = "delete from dotBrgyData where NOS=@id";
                            cmd.Parameters.AddWithValue("@id", trackId.Text);
                            cmd.ExecuteNonQuery();
                            DisplayAll();
                            Clear(this);
                        }
                    }
                }
            }
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

                        SaveResidentImage();

                        string fileN = @"C:\DotBrgy\Images\residentImage.png";
                        //Initialize a file stream to read the image file
                        FileStream fs = new FileStream(fileN, FileMode.Open, FileAccess.Read);

                        //Initialize a byte array with size of stream
                        byte[] imgByteArr = new byte[fs.Length];

                        //Read data from the file stream and put into the byte array
                        fs.Read(imgByteArr, 0, Convert.ToInt32(fs.Length));

                        //Close a file stream
                        fs.Close();

                        WaitCursor();
                        con.Open();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "update dotBrgyData set img=@img,brgyID=@brgyID,firstname=@firstname,middlename=@middlename,lastname=@lastname,purok=@purok,Sex=@Sex,Birthdate=@Birthdate,BirthPlace=@BirthPlace,CivilStatus=@CivilStatus,EducationalAttainment=@EducationalAttainment," +
                            "Philhealth=@Philhealth,RegisteredVoter=@RegisteredVoter,Occupation=@Occupation,FamilyPlanning=@FamilyPlanning,Faucet=@Faucet,ComfortRoom=@ComfortRoom,SeniorID=@SeniorID,PWD=@PWD,IndigenousPeople=@IndigenousPeople,Membership=@Membership where NOS=" + trackId.Text;
                        //cmd.CommandText = "update transactions  set ownerFirstname=@firstname,ownerMiddlename=@middlename,ownerLastname=@lastname,CivilStatus=@CivilStatus,succeedingAction=@action where id=" + TrackId.Text;

                        cmd.Parameters.AddWithValue("img", imgByteArr);

                        cmd.Parameters.AddWithValue("@brgyID", brgyID.Text);
                        cmd.Parameters.AddWithValue("@firstname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(firstname.Text));
                        cmd.Parameters.AddWithValue("@middlename", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(middlename.Text));
                        cmd.Parameters.AddWithValue("@lastname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(lastname.Text));

                       
                        cmd.Parameters.AddWithValue("@purok", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(purok.Text));
                        cmd.Parameters.AddWithValue("@Sex", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(Sex.Text));
                        cmd.Parameters.AddWithValue("Birthdate", Birthdate.Text);
                        cmd.Parameters.AddWithValue("BirthPlace", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(birthPlace.Text));
                        cmd.Parameters.AddWithValue("@CivilStatus", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(CivilStatus.Text));


                        cmd.Parameters.AddWithValue("@EducationalAttainment", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(educationAttainment.Text));
                        cmd.Parameters.AddWithValue("@Philhealth", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(philHealth.Text));
                        cmd.Parameters.AddWithValue("@RegisteredVoter", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(registeredVoter.Text));
                        cmd.Parameters.AddWithValue("@Occupation", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(occupation.Text));
                        cmd.Parameters.AddWithValue("@FamilyPlanning", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(familyPlanning.Text));
                        cmd.Parameters.AddWithValue("@Faucet", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(faucet.Text));
                        cmd.Parameters.AddWithValue("@ComfortRoom", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(comfortRoom.Text));
                        cmd.Parameters.AddWithValue("@SeniorID", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(seniorID.Text));
                        cmd.Parameters.AddWithValue("@PWD", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(pwd.Text));
                        cmd.Parameters.AddWithValue("@IndigenousPeople", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(indigenousPeople.Text));
                        cmd.Parameters.AddWithValue("@Membership", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(membership.Text));
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Record has been successfully updated!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Question);
                        DisplayAll();
                        //display_transaction();
                        Clear(this);
                        NormalCursor();

                        /*
                        using (SQLiteConnection conStat = new SQLiteConnection(sqliteConnectionString))
                        {
                            using (SQLiteCommand cmdStat = conStat.CreateCommand())
                            {
                                WaitCursor();
                                conStat.Open();
                                cmdStat.CommandType = CommandType.Text;
                                cmdStat.CommandText = "update statistic set purok=@p where NOS=" + statNumber.Text;

                                cmdStat.Parameters.AddWithValue("@p", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(sample.Text));

                                cmdStat.ExecuteNonQuery();

                                MessageBox.Show("Record has been successfully updated!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Question);
                                DisplayAll();
                                Clear(this);
                                NormalCursor();
                            }
                        } */
                    }
                }
            }
        }

        private void ButtoGoAdmin_Click(object sender, RoutedEventArgs e)
        {
            LoginWindow open = new LoginWindow();
            open.Show();
            this.Hide();
        }

        private void countPurok_TextChanged(object sender, TextChangedEventArgs e)
        {

            /*
      using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
      {
          con.Open();
          using (SQLiteCommand command = new SQLiteCommand("Select count (*) from statistic where purok = '" + purok.Text.Trim() + "'", con))
          {
              var result = command.ExecuteScalar();
              int i = Convert.ToInt32(result);
              if (i != 0)
              {
                  return;
              }
              else
              {
                  statNumber.Text = string.Empty;
              }
          }
      }*/
        }

        private void refreshStat_Click(object sender, RoutedEventArgs e)
        {
            Clear(this);
            DisplayAll();
        }                             

        private void SearchHistory_TextChanged(object sender, TextChangedEventArgs e)
        {
            //documents.Visibility = Visibility.Collapsed;
            Regex rx = new Regex("^[a-zA-Z]+$");
            if (rx.IsMatch(SearchHistory.Text))
            {
                DataView dv = dbHistory.ItemsSource as DataView;
                dv.RowFilter = string.Format("lastname LIKE '%{0}%'", SearchHistory.Text);
            }
            else
            {
                return;
            }

        }

        private void DeleteAllHistory_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dbHistory == null || dbHistory.Items.Count == 0)
                {
                    MessageBox.Show(this, "No data found!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    MessageBoxResult delete_all = MessageBox.Show(this, "Delete History?", "DotBrgy", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
                    if (delete_all == MessageBoxResult.Yes)
                    {

                        using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                        {
                            using (SQLiteCommand deleteHistory = con.CreateCommand())
                            {
                                using (SQLiteCommand cleanHistory = con.CreateCommand())
                                {
                                    con.Open();

                                    deleteHistory.CommandText = "delete from history";
                                    deleteHistory.ExecuteNonQuery();

                                    cleanHistory.CommandText = "DELETE FROM sqlite_sequence WHERE name = 'history'";
                                    cleanHistory.CommandText = "UPDATE sqlite_sequence SET seq = 0 WHERE name = 'history'";
                                    cleanHistory.ExecuteNonQuery();
                                    DisplayHistory();
                                    MessageBox.Show("No history found!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);


                                    //Clean();
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

        private void printHistory_Click(object sender, RoutedEventArgs e)
        {
            if (dbHistory.Items.Count == 0)
            {
                MessageBox.Show("No record found!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else
            {
                MessageBoxResult result = MessageBox.Show("You are about to generate all your history, Proceed?", "DotBrgy", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.No)
                {
                    Clear(this);
                    return;
                }
                else
                {
                    WaitCursor();
                    PrintHistory();
                    NormalCursor();
                }
            }
        }

        private void PrintHistory()
        {
            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                wb = excel.Workbooks.Add();
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;


                for (int Idx = 0; Idx < dbHistory.Columns.Count; Idx++)
                {
                    ws.Range["A1"].Offset[0, Idx].Value = dbHistory.Columns[Idx].Header;
                }

                for (int rowIndex = 0; rowIndex < dbHistory.Items.Count; rowIndex++)
                {
                    for (int columnIndex = 0; columnIndex < dbHistory.Columns.Count; columnIndex++)
                    {
                        ws.Range["A2"].Offset[rowIndex, columnIndex].Value = (dbHistory.Items[rowIndex] as DataRowView).Row.ItemArray[columnIndex].ToString();
                    }
                    excel.Columns.AutoFit();
                    excel.Rows.AutoFit();
                }
                for (int i = 1; i < dbHistory.Items.Count - 1; i++)
                {

                }
                MessageBox.Show("Thank you for your patience, Click OK to view your files!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                excel.Visible = true;
            }
            catch (COMException ex)
            {
                MessageBox.Show("Error accessing Excel: " + ex.ToString(), "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.ToString(), "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
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

        private void transType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (transType.SelectedValue)
            {
                case "Business":
                    businessName.Visibility = Visibility.Visible;
                    businessName.Focus();
                    break;
                default:
                    businessName.Visibility = Visibility.Collapsed;
                    break;
            }
        }

        private void printAll_Click(object sender, RoutedEventArgs e)
        {
            if (dbData.Items.Count == 0)
            {
                MessageBox.Show("No record found!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else if (trackId.Text == string.Empty)
            {
                MessageBoxResult result = MessageBox.Show("You are about to generate all data in the table, Proceed?", "DotBrgy", MessageBoxButton.YesNo, MessageBoxImage.Information);
                if (result == MessageBoxResult.No)
                {
                    return;
                }
                else
                {
                    WaitCursor();
                    try
                    {
                        excel = new Microsoft.Office.Interop.Excel.Application();
                        wb = excel.Workbooks.Add();
                        ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;


                        for (int Idx = 0; Idx < dbData.Columns.Count; Idx++)
                        {
                            ws.Range["A1"].Offset[0, Idx].Value = dbData.Columns[Idx].Header;
                        }

                        for (int rowIndex = 0; rowIndex < dbData.Items.Count; rowIndex++)
                        {
                            for (int columnIndex = 0; columnIndex < dbData.Columns.Count; columnIndex++)
                            {
                                ws.Range["A2"].Offset[rowIndex, columnIndex].Value = (dbData.Items[rowIndex] as DataRowView).Row.ItemArray[columnIndex].ToString();
                            }
                            excel.Columns.AutoFit();
                            excel.Rows.AutoFit();
                        }
                        MessageBox.Show("Thank you for your patience, Click OK to view your files!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        excel.Visible = true;
                    }
                    catch (COMException ex)
                    {
                        MessageBox.Show("Error accessing Excel: " + ex.ToString(), "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.ToString(), "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    NormalCursor();
                }
            }
        }

        void Tracker()
        {
            try
            {
                using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                {

                    using (SQLiteCommand cmd = con.CreateCommand())
                    {
                        //string var;
                        //var = "Track number " + generateId.Text;
                        //int monthParse = month.SelectedIndex + 1;
                        con.Open();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "insert into history(residentCode,dateTime,firstname,middlename,lastname,documents,purok,Sex,Birthdate)" +
                        " values(@residentCode,@dateTime,@firstname,@middlename,@lastname,@documents,@purok,@Sex,@Birthdate)";

                        cmd.Parameters.AddWithValue("@residentCode", code.Text);
                        cmd.Parameters.AddWithValue("@dateTime", DateTime.Now.ToString());
                        cmd.Parameters.AddWithValue("@firstname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(firstname.Text));
                        cmd.Parameters.AddWithValue("@middlename", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(middlename.Text));
                        cmd.Parameters.AddWithValue("@lastname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(lastname.Text));

                        cmd.Parameters.AddWithValue("@purok", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(purok.Text));
                        cmd.Parameters.AddWithValue("@Sex", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(Sex.Text));
                        cmd.Parameters.AddWithValue("@Birthdate", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(Birthdate.Text));

                        if (transType.Text == string.Empty)
                        {
                            cmd.Parameters.AddWithValue("@documents", "Brgy ID");
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@documents", transType.Text);
                        }
                        cmd.ExecuteNonQuery();
                        DisplayHistory();
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void printDocs_Click(object sender, RoutedEventArgs e)
        {
            if (dbData.Items.Count == 0)
            {
                MessageBox.Show("No record found!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else if (trackId.Text == string.Empty)
            {
                MessageBox.Show("Please select valid item!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else if (transType.Text == string.Empty)
            {
                MessageBoxResult result = MessageBox.Show("Please select transacton documents!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                transType.IsDropDownOpen = true;
            }
            else
            {
                Tracker();
                switch (transType.Text)
                {
                    case "Business":
                        WaitCursor();
                        SaveQRCODEImage();
                        CreateWordDocument(@"C:\DotBrgy\Documents\business.docx", @"C:\DotBrgy\Print\print.docx");
                        NormalCursor();
                        break;
                    case "Clearance":
                        WaitCursor();
                        SaveQRCODEImage();
                        CreateWordDocument(@"C:\DotBrgy\Documents\clearance.docx", @"C:\DotBrgy\Print\print.docx");
                        NormalCursor();
                        break;
                    case "Indigency":
                        WaitCursor();
                        SaveQRCODEImage();
                        CreateWordDocument(@"C:\DotBrgy\Documents\indigency.docx", @"C:\DotBrgy\Print\print.docx");
                        NormalCursor();
                        break;
                    case "Residency":
                        WaitCursor();
                        SaveQRCODEImage();
                        CreateWordDocument(@"C:\DotBrgy\Documents\residency.docx", @"C:\DotBrgy\Print\print.docx");
                        NormalCursor();
                        break;
                }
            }
        }

        //Find and Replace Method
        private void frBrgyID(Microsoft.Office.Interop.Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            //object read_only = false;
            //object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

        //Create the Doc Method
        private void CreateBrgyID(object filename, object SaveAs)
        {
            try
            {

                wordApp = new Microsoft.Office.Interop.Word.Application();
                object missing = Missing.Value;
                Microsoft.Office.Interop.Word.Document myWordDoc = null;

                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;



                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing,
                                           ref readOnly, ref missing, ref missing,
                                           ref missing, ref missing, ref missing,
                                           ref missing, ref missing, ref missing,
                                           ref missing, ref missing, ref missing,
                                           ref missing, ref missing);
                    myWordDoc.Activate();

                    //find and replace
                    this.frBrgyID(wordApp, "<firstname>", firstname.Text);
                    this.frBrgyID(wordApp, "<middlename>", middlename.Text);
                    this.frBrgyID(wordApp, "<lastname>", lastname.Text);
                    this.frBrgyID(wordApp, "<age>", age.Text);
                    this.frBrgyID(wordApp, "<CivilStatus>", CivilStatus.Text);
                    this.frBrgyID(wordApp, "<purok>", purok.Text);
                    this.frBrgyID(wordApp, "<date>", DateTime.Now.ToString("MMMM dd, yyyy"));
                    this.frBrgyID(wordApp, "<day>", DateTime.Now.ToString("dd"));
                    this.frBrgyID(wordApp, "<month>", DateTime.Now.ToString("MMMM"));
                    this.frBrgyID(wordApp, "<year>", DateTime.Now.ToString("yyyy"));
                    this.frBrgyID(wordApp, "<Sex>", Sex.Text);
                    this.frBrgyID(wordApp, "<Birthdate>", Birthdate.Text);
                    //this.FindAndReplace(wordApp, "<or>", DateTime.Now.ToString("yyyyMMdd-HHmmss-fff"));
                    this.FindAndReplace(wordApp, "<or>", code.Text);
                    this.frBrgyID(wordApp, "<business>", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(businessName.Text));
                    //this.FindAndReplace(wordApp, "<date>", DateTime.Now.ToShortDateString());
                }
                else
                {
                    MessageBox.Show("File not Found!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                //Save as
                myWordDoc.SaveAs(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                //insert image into document
                //myWordDoc.Bookmarks["add"].Range.InsertParagraph.myWordDoc

                //myWordDoc.Bookmarks["add"].Range.InsertParagraph.myWordDoc;

                //var shape = myWordDoc.Shapes.AddPicture(@"C:\DotBrgy\Images\eBrgyQR.png", false, true);

                //shape.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //shape.Width = 100;
                //shape.Height = 100;
                /*   
                var outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
                var logoimage = Path.Combine(outPutDirectory, @"C:\DotBrgy\Images\eBrgyQR.png");

                var keyword = "LOGO";
                var sel = wordApp.Selection;
                sel.Find.Text = string.Format("[{0}]", keyword);
                wordApp.Selection.Find.Execute(keyword);

                Microsoft.Office.Interop.Word.Range range = wordApp.Selection.Range;
                if (range.Text.Contains(keyword))
                {

                    //gets desired range here it gets last character to make superscript in range 
                    Microsoft.Office.Interop.Word.Range temprange = myWordDoc.Range(range.End - 4, range.End);//keyword is of 4 charecter range.End - 4
                    temprange.Select();

                    Microsoft.Office.Interop.Word.Selection currentSelection = wordApp.Selection;
                    //currentSelection.Font.Superscript = 1;

                    sel.Find.Execute(Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne);
                    sel.Range.Select();
                    var imagePath = Path.GetFullPath(string.Format(logoimage, keyword));
                    var shape = sel.InlineShapes.AddPicture(FileName: imagePath, LinkToFile: false, SaveWithDocument: true).ConvertToShape();
                    shape.Width = 40;
                    shape.Height = 40;
                    shape.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapInline;
                }  
                */
                var outPutDirectory1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
                var logoimage1 = Path.Combine(outPutDirectory1, @"C:\DotBrgy\Images\residentImage.png");

                var keyword1 = "OKAY";
                var sel1 = wordApp.Selection;
                sel1.Find.Text = string.Format("[{0}]", keyword1);
                wordApp.Selection.Find.Execute(keyword1);

                Microsoft.Office.Interop.Word.Range range1 = wordApp.Selection.Range;


                //gets desired range here it gets last character to make superscript in range 
                Microsoft.Office.Interop.Word.Range temprange1 = myWordDoc.Range(range1.End - 4, range1.End);//keyword is of 4 charecter range.End - 4
                temprange1.Select();

                //Microsoft.Office.Interop.Word.Selection currentSelection = wordApp.Selection;
                //currentSelection.Font.Superscript = 1;

                sel1.Find.Execute(Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne);
                sel1.Range.Select();
                var imagePath1 = Path.GetFullPath(string.Format(logoimage1, keyword1));
                var shape1 = sel1.InlineShapes.AddPicture(FileName: imagePath1, LinkToFile: false, SaveWithDocument: true).ConvertToShape();
                shape1.Width = 50;
                shape1.Height = 50;
                shape1.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapInline;




                var outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
                var logoimage = Path.Combine(outPutDirectory, @"C:\DotBrgy\Images\eBrgyQR.png");

                var keyword = "LOGO";
                var sel = wordApp.Selection;
                sel.Find.Text = string.Format("[{0}]", keyword);
                wordApp.Selection.Find.Execute(keyword);

                Microsoft.Office.Interop.Word.Range range = wordApp.Selection.Range;

                //gets desired range here it gets last character to make superscript in range 
                Microsoft.Office.Interop.Word.Range temprange = myWordDoc.Range(range.End - 4, range.End);//keyword is of 4 charecter range.End - 4
                temprange.Select();

                Microsoft.Office.Interop.Word.Selection currentSelection = wordApp.Selection;
                //currentSelection.Font.Superscript = 1;

                sel.Find.Execute(Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne);
                sel.Range.Select();
                var imagePath = Path.GetFullPath(string.Format(logoimage, keyword));
                var shape = sel.InlineShapes.AddPicture(FileName: imagePath, LinkToFile: false, SaveWithDocument: true).ConvertToShape();
                shape.Width = 40;
                shape.Height = 40;
                shape.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapInline;

                myWordDoc.Close();
                wordApp.Quit();


                MessageBoxResult result = MessageBox.Show("File successfully saved in your disk!, Click yes to Open.", "DotBrgy", MessageBoxButton.YesNo, MessageBoxImage.Asterisk);
                if (result == MessageBoxResult.Yes)
                {
                    /*
                    System.Diagnostics.Process myProcess = new System.Diagnostics.Process();
                    myProcess.StartInfo.FileName = @"C:\Backup\print.docx";
                    myProcess.StartInfo.CreateNoWindow = true;
                    myProcess.Start();
                    */

                    System.Diagnostics.Process.Start(@"C:\DotBrgy\Print\brgyID-print.docx");
                }
                else
                {
                    //KillWord();
                    //MessageBox.Show("File successfully saved in your disk!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    return;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error: " + ex.ToString(), "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Error);
                MessageBox.Show("Word cannot save this file because it is already open elsewhere.", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Error);
                KillWord();
            }
        }

        private void printID_Click(object sender, RoutedEventArgs e)
        {
            if (trackId.Text == string.Empty)
            {
                MessageBox.Show("Please select valid item!", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else
            {
                WaitCursor();
                Tracker();
                SaveResidentImage();
                CreateBrgyID(@"C:\DotBrgy\ID\brgyID-format.docx", @"C:\DotBrgy\Print\brgyID-print.docx");
                NormalCursor();
            }
        }

        private void resCodeHistoryRadio_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                DataView dv = dbHistory.ItemsSource as DataView;
                dv.RowFilter = "Convert(residentCode, 'System.String') like '%" + SearchHistory.Text + "%'"; //where n is a column name of the DataTable
            }
            catch (System.Exception)
            {
                MessageBox.Show("No data found", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void dateHistoryRadio_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                DataView dv = dbHistory.ItemsSource as DataView;
                dv.RowFilter = "Convert(dateTime, 'System.String') like '%" + SearchHistory.Text + "%'"; //where n is a column name of the DataTable
            }
            catch (System.Exception)
            {
                MessageBox.Show("No data found", "DotBrgy", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private void refreshTracker_Click(object sender, RoutedEventArgs e)
        {
            DisplayHistory();
            Clear(this);
        }

        private void startCamera_Click(object sender, RoutedEventArgs e)
        {
            LocalWebCam.Start();
            cameras.Visibility = Visibility.Visible;
            captureCamera.Visibility = Visibility.Visible;
            stopCamera.Visibility = Visibility.Visible;
            startCamera.Visibility = Visibility.Collapsed;
        }

        private void captureCamera_Click(object sender, RoutedEventArgs e)
        {
            //Clear(this);
            LocalWebCam.Stop();
            captureCamera.Visibility = Visibility.Collapsed;
            startCamera.Visibility = Visibility.Visible;
        }

        private void stopCamera_Click(object sender, RoutedEventArgs e)
        {
            if (uploadImage.Source != null)
            {
                startCamera.Visibility = Visibility.Visible;
                LocalWebCam.Stop();
                stopCamera.Visibility = Visibility.Collapsed;
                cameras.Visibility = Visibility.Collapsed;
                uploadImage.Source = null;
                captureCamera.Visibility = Visibility.Collapsed;
            }
            else
            {
                startCamera.Visibility = Visibility.Visible;
            }
        }

        private void lowest_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lowest.Text.ToString() == String.Empty)
            {
                highest.IsEnabled = false;
            }
            else
            {
                highest.IsEnabled = true;
            }
        }

        private void highest_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (highest.Text.ToString() == String.Empty)
            {
                return;
            }
            else
            {
                DataView dv = dbData.ItemsSource as DataView;
                dv.RowFilter = "Convert(age, 'System.String') >= '" + int.Parse(lowest.Text.ToString()) + "' and  Convert(age, 'System.String') <= '" + int.Parse(highest.Text.ToString()) + "'"; //where n is a column name of the DataTable 
            }
        }

        private void openRight_Click(object sender, RoutedEventArgs e)
        {
            Main.IsEnabled = false;
        }

        private void closeRight_Click(object sender, RoutedEventArgs e)
        {
            Main.IsEnabled=true;
        }

        private void openLeft_Click(object sender, RoutedEventArgs e)
        {
            Main.IsEnabled = false;
        }

        private void closeLeft_Click(object sender, RoutedEventArgs e)
        {
            Main.IsEnabled = true;
        }
    }
}