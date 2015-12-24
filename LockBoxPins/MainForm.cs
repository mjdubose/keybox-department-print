using System;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Drawing.Drawing2D;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;



namespace LockBoxPins
{
    public partial class MainForm : Form
    {
        private PrintPreviewDialog Preview = new PrintPreviewDialog();
        //Initialize these as public because it needs to be called by more than one Method.
        private PrintDocument printDocument1 = new PrintDocument();
        public DataGridView DataGrid1 = new DataGridView();
       // public int PrintRow;

        // string to hold the entire string of things to be printed.
        public string ToBePrinted = string.Empty;

        //a variable to hold the portion of the document that is not printed.
        public string DocumentContents = string.Empty;
        public MainForm()
        {
            InitializeComponent();

            Text = "Lock Box Pins";

            //Initialize Controls
            GroupBox DataGroup = new GroupBox();
            Button Importbttn = new Button();
            Button Exitbttn = new Button();
            Button Printbttn = new Button();
            Button Aboutbttn = new Button();

            //Add Properties to Import Button
            Importbttn.Location = new Point(10, 50);
            Importbttn.Size = new Size(80, 40);
            Importbttn.Text = "Click to Import";
            Importbttn.Click += new EventHandler(OnClickImport);

            Aboutbttn.Location = new Point(510, 50);
            Aboutbttn.Size = new Size(80, 40);
            Aboutbttn.Text = "About LockBoxPins";
            Aboutbttn.Click += new EventHandler(OnClickAbout);

            //Add properties to Exit Button
            Exitbttn.Location = new Point(600, 50);
            Exitbttn.Size = new Size(80, 40);
            Exitbttn.Text = "Exit";
            Exitbttn.Click += new EventHandler(OnClickExit);

            Printbttn.Location = new Point(100, 50);
            Printbttn.Size = new Size(80, 40);
            Printbttn.Text = "Print Pins";
            Printbttn.Click += new EventHandler(OnClickPrintPreview);

            //Add Properties to Group Box
            DataGroup.Location = new Point(10, 100);
            DataGroup.Size = new Size(705, 450);

            //Add Properties to Data Grid View
            DataGrid1.Location = new Point(10, 10);
            DataGrid1.Size = new Size(690, 435);
            DataGrid1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.ColumnHeader;
            DataGrid1.AllowUserToAddRows = false;


            printDocument1.PrintPage += new PrintPageEventHandler(printDocument1_PrintPage);
            //Add Datagrid1 to Dataroup
            DataGroup.Controls.Add(DataGrid1);

            //Add Controls to MainForm
            Controls.Add(Importbttn);
            Controls.Add(Printbttn);
            Controls.Add(Aboutbttn);
            Controls.Add(Exitbttn);
            Controls.Add(DataGroup);
        }

        private void OnClickExit(object sender, System.EventArgs e)
        {
            Close();
        }

        private void OnClickAbout(object sender, System.EventArgs e)
        {
            AboutBox1 Aboutbox1 = new AboutBox1();
            Aboutbox1.Show();
        }

        private void OnClickImport(object sender, System.EventArgs e)
        {
            //Initialize Excel Components.
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Application xlApp = new Excel.Application();

            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "Excel Files(.xls)|*.xls| Excel Files(.xlsx)|*.xlsx| Excel Files(*.xlsm)|*.xlsm";
            of.Title = "Open Excel File to Work with";

            //If Excel file is opened.
            if (of.ShowDialog() == DialogResult.OK)
            {
                //MessageBox.Show("Import Data");
                xlWorkBook = xlApp.Workbooks.Open(of.FileName,
                  Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //Shows the first sheet
                xlWorkSheet.Activate();

                //Grabs the worksheet name this is used in the OLEDB connection below.
                string WorkSheetName = xlWorkSheet.Name.ToString();
                xlWorkBook.Close(Type.Missing, Type.Missing, Type.Missing);
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);                
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                //Excel application is visible if true and not visible if false.  True for testing purposes.
                // xlApp.Visible = true;

                string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + of.FileName + ";Extended Properties='Excel 12.0;EMEX=1;'";
                using (OleDbConnection Conn = new OleDbConnection(connStr))
                {
                    Conn.Open();
                    string Query = string.Format("SELECT * FROM [{0}]", WorkSheetName + "$");

                    using (OleDbDataAdapter TA = new OleDbDataAdapter(Query, Conn))
                    {
                        DataSet DS = new DataSet();
                        DataTable dt = new DataTable();
                        TA.Fill(DS);
                        DataGrid1.DataSource = DS.Tables[0];
                    }
                    Conn.Close();
                }
            }

        }

        private void PrepareDocument()
        {
            ToBePrinted = string.Empty;
            for (int i = 0; i < DataGrid1.RowCount; i++)
            {
                var row = DataGrid1.Rows[i];

                var builder = new StringBuilder();

                var thelist = row.Cells.Cast<DataGridViewCell>().ToList();

             

                builder.Append(thelist[7].Value.ToString() + Environment.NewLine);
                builder.AppendLine(string.Format("{0,-25} Your Pin Is : {1}.",thelist[1].Value.ToString() + " " + thelist[0].Value.ToString(),thelist[4].Value.ToString()));
                
                ToBePrinted = ToBePrinted + builder.ToString();
            }
            DocumentContents = ToBePrinted;
        }
        private void OnClickPrintPreview(object sender, EventArgs e)
        {

            if (DataGrid1.DataSource != null)
            {
                PrepareDocument();
                Preview.Document = printDocument1;
                Preview.ShowDialog();
            }
            else
            {
                MessageBox.Show("Please Import your data before printing");
            }
        }

        //found original code here on stackoverlow.com
        //http://stackoverflow.com/questions/27448856/how-to-print-the-values-of-datagridview-in-c



        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            int ypos = 1;
            float pageheight = e.PageBounds.Height;
            string deptslots = string.Empty;
            string picturename = "Greenlit.jpg";
            PrintImageCentered(e, ref ypos,"logo_black.gif");
            PrintUntilNewLine(e, ref ypos, ref deptslots, ref picturename);
            PrintUntilNewLine(e, ref ypos, ref deptslots, ref picturename);

            PrintImageCentered(e, ref ypos, "PINENTRY.png");

            PrintBoxUsageInformation(e, ref ypos, "Key Box Usage:" + Environment.NewLine );
            PrintBoxUsageInformation(e,ref ypos," \tAt the prompt, enter your pin." + Environment.NewLine +"\tOpen the door using the latch on the right hand side." + Environment.NewLine +"\tThe key positions available to you will be lit up in green.");

            PrintImage(e, ref ypos, picturename);      

           

            PrintBoxUsageInformation(e, ref ypos, deptslots);

            PrintBoxUsageInformation(e, ref ypos, Environment.NewLine +"If the box times out while you are taking out a key, or putting the key back, re-enter your pin"+Environment.NewLine );
            PrintBoxUsageInformation(e, ref ypos, "and continue."+ Environment.NewLine+Environment.NewLine+"When you enter your pin to put a key back, only the slot where you pulled the key will be");
            PrintBoxUsageInformation(e, ref ypos, "lit up."+Environment.NewLine+Environment.NewLine+"The box is set to allow you eight hours with a set of keys.  If you are working over, you'll need");
            PrintBoxUsageInformation(e, ref ypos, "to return your key and repull the key before eight hours are up to reset the timer.");
            while (ypos + 60 < pageheight)
            {
                ypos = ypos + 60;
            }


            if (ypos + 60 > pageheight && DocumentContents.Length > 0)
            {
                e.HasMorePages = true;
            }
            else
            {
                e.HasMorePages = false;
                DocumentContents = ToBePrinted;
            }

        }
      private void PrintImage(PrintPageEventArgs e, ref int ypos, string filename)
        {
            float pagewidth = e.PageBounds.Width;
            Image img1 = Image.FromFile(filename);
            Image img = FixedSize(img1, 400, 200);
            Point loc = new Point((int)((e.PageBounds.Width / 2) - (img.Width / 1.5)), ypos);
            ypos = ypos + (int)(img.Height * 1.5);
            e.Graphics.DrawImage(img, loc);
            img.Dispose();


        }

        private void PrintImageCentered(PrintPageEventArgs e, ref int ypos, string filename)
        {
            float pagewidth = e.PageBounds.Width;
            Image img = Image.FromFile(filename);

            
            Point loc = new Point((int)(((pagewidth - img.Width)/2)), ypos);
            ypos = ypos + img.Height;
            e.Graphics.DrawImage(img, loc);
            img.Dispose();

           
        }
        private void PrintUntilNewLine(PrintPageEventArgs e, ref int ypos,ref string dept, ref string picturename)
        {
            string x = DocumentContents.Substring(0, DocumentContents.IndexOf(Environment.NewLine));
            switch (x){

                case "Engineering Electricians":
                    dept = "Your available key pull positions are 41-48.";
                    picturename = "Engineering Electricians.jpg";
                        break;
                case "Engineering Carpet and Tile":
                        dept = "Your available key pull positions are 50-55.";
                        picturename = "Engineering Carpet and Tile.jpg";
                    break;
                case "Engineering HVAC":
                    dept = "Your available key pull positions are 64-75.";
                    picturename = "Engineering HVAC.jpg";
                    break;
                case "Engineering Comm Appliance":
                    dept = "Your available key pull positions are 26-30.";
                    picturename = "Engineering Comm Appliance.jpg";
                    break;
                case "Engineering Painters":
                    dept = "Your available key pull positions are 13-24.";
                    picturename = "Engineering Painters.jpg";
                    break;
                case "Engineering Carpenters":
                    dept = "Your available key pull positions are 32-39.";
                    picturename = "Engineering Carpenters.jpg";
                    break;
                case "Engineering Plumber":
                    dept = "Your available key pull positions are 77-81.";
                    picturename = "Engineering Plumber.jpg";
                    break;
                case "Engineering Stationary":
                    dept = "Your available key pull positions are 83-86.";
                    picturename = "Engineering Stationary.jpg";
                    break;
                case "Engineering General Mechanics":
                    dept = "Your available key pull positions are 57-62.";
                    picturename = "Engineering General Mechanics.jpg";
                    break;
                 default:
                  
                     break ;
                    
            }
            var currentstring = e.Graphics.MeasureString(x + Environment.NewLine +" ", new Font("Arial", 14.0f));
            e.Graphics.DrawString(x + Environment.NewLine +" ", new Font("Arial", 14.0f), Brushes.Black, 1, ypos);
            ypos = ypos + (int) currentstring.Height;
            DocumentContents = RemovePrintedString(DocumentContents, x + Environment.NewLine);
        }
        static Image FixedSize(Image imgPhoto, int Width, int Height)
        {
            int sourceWidth = imgPhoto.Width;
            int sourceHeight = imgPhoto.Height;
            int sourceX = 0;
            int sourceY = 0;
            int destX = 0;
            int destY = 0;

            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;

            nPercentW = ((float)Width / (float)sourceWidth);
            nPercentH = ((float)Height / (float)sourceHeight);
            if (nPercentH < nPercentW)
            {
                nPercent = nPercentH;
                destX = Convert.ToInt16((Width -
                              (sourceWidth * nPercent)) / 2);
            }
            else
            {
                nPercent = nPercentW;
                destY = Convert.ToInt16((Height -
                              (sourceHeight * nPercent)) / 2);
            }

            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);

            Bitmap bmPhoto = new Bitmap(Width, Height,
                              PixelFormat.Format24bppRgb);
            bmPhoto.SetResolution(imgPhoto.HorizontalResolution,
                             imgPhoto.VerticalResolution);

            Graphics grPhoto = Graphics.FromImage(bmPhoto);
            grPhoto.Clear(Color.White);
            grPhoto.InterpolationMode =
                    InterpolationMode.HighQualityBicubic;

            grPhoto.DrawImage(imgPhoto,
                new Rectangle(destX, destY, destWidth, destHeight),
                new Rectangle(sourceX, sourceY, sourceWidth, sourceHeight),
                GraphicsUnit.Pixel);

            grPhoto.Dispose();
            return bmPhoto;
        }

        private void PrintBoxUsageInformation(PrintPageEventArgs e, ref int ypos, string MessageToBePrinted)
        {
           var currentstring = e.Graphics.MeasureString(MessageToBePrinted, new Font("Arial",14.0f));
            e.Graphics.DrawString(MessageToBePrinted, new Font("Arial", 14.0f), Brushes.Black, 1, ypos);
               ypos = ypos + (int)currentstring.Height;
        }
        private string RemovePrintedString(string string1, string string2)
        {
            string string1_part1 = string1.Substring(0, string1.IndexOf(string2));
            string string1_part2 = string1.Substring(
                string1.IndexOf(string2) + string2.Length, string1.Length - (string1.IndexOf(string2) + string2.Length));
            return string1_part1 + string1_part2;
        }
    }
}
