using System;
using System.Globalization;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using PdfFileAnalyzer;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;
using System.Data.OleDb;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        byte[] StreamByteArray;
        String ConfigPath = AppDomain.CurrentDomain.BaseDirectory + "Config" + "\\Book1.xlsx";
        Dictionary<string, string> MyDict = new Dictionary<string, string>();
        DataSet dataSet;
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            
            String InputPath = AppDomain.CurrentDomain.BaseDirectory + "Input";
            String OutputPath = AppDomain.CurrentDomain.BaseDirectory + "Output";

            if (!Directory.Exists(InputPath))
            {
                Directory.CreateDirectory(InputPath);
            }
            if (!Directory.Exists(OutputPath))
            {
                Directory.CreateDirectory(OutputPath);
            }

            string[] InputPaths = Directory.GetFiles(InputPath);
            foreach (string filePath in InputPaths)
            {
                FileAttributes attrs = File.GetAttributes(filePath);
                if (attrs.HasFlag(FileAttributes.ReadOnly))
                    File.SetAttributes(filePath, attrs & ~FileAttributes.ReadOnly);
                    File.Delete(filePath);
            }
            string[] OutputPaths = Directory.GetFiles(OutputPath);
            foreach (string filePath in OutputPaths)
            {
                FileAttributes attrs = File.GetAttributes(filePath);
                if (attrs.HasFlag(FileAttributes.ReadOnly))
                    File.SetAttributes(filePath, attrs & ~FileAttributes.ReadOnly);
                    File.Delete(filePath);
            }

            openFileDialog2.Filter = "HPT files | *.hpt";
            openFileDialog2.Title = "Select HPT file";
            openFileDialog2.Multiselect = true;

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                
                foreach (string fileName in openFileDialog2.FileNames)
                {
                    try
                    {
                        File.Copy(fileName, Path.ChangeExtension(InputPath + @"\" + System.IO.Path.GetFileName(fileName), ".pdf"),true);

                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                    }
                    
                }

                string[] FiletoProcessPaths = Directory.GetFiles(InputPath);
                foreach (string filePath in FiletoProcessPaths)
                {
                Parse(filePath);
                }
            }
            if (MessageBox.Show("Conversion is complete." + Environment.NewLine + "Open file location?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(AppDomain.CurrentDomain.BaseDirectory + "Output");
            }

        }
        

        private void CreateOutput(string PDFInput)
        {
            try
            {
                
                iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(@PDFInput);
                FileStream fs = new FileStream(@PDFInput.Replace("Input","Output"), FileMode.Create, FileAccess.Write);
                PdfStamper stamp = new PdfStamper(reader, fs);
                //Loop from here

                //MessageBox.Show(dataGridView1.RowCount.ToString());
                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    string comptype = dataGridView1.Rows[i].Cells[3].Value.ToString();

                    string complocation = dataGridView1.Rows[i].Cells[1].Value.ToString();

                    float llx = float.Parse(complocation.Split(' ')[0]);
                    float lly = float.Parse(complocation.Split(' ')[1]);
                    float urx = float.Parse(complocation.Split(' ')[2]);
                    float ury = float.Parse(complocation.Split(' ')[3]);


                    if (comptype == "TE")
                    {
                        int x = Int32.Parse(dataGridView1.Rows[i].Cells[2].Value.ToString());
                        PdfFormField ff = PdfFormField.CreateTextField(stamp.Writer, false, false, 50);
                        ff.SetWidget(new iTextSharp.text.Rectangle(llx, lly, urx, ury), PdfAnnotation.HIGHLIGHT_INVERT);
                        ff.SetFieldFlags(PdfAnnotation.FLAGS_PRINT);
                        ff.FieldName = dataGridView1.Rows[i].Cells[0].Value.ToString();
                        stamp.AddAnnotation(ff, x);
                    }
                    else if (comptype == "CB")
                    {
                        int x = Int32.Parse(dataGridView1.Rows[i].Cells[2].Value.ToString());
                        RadioCheckField fCell = new RadioCheckField(stamp.Writer, new iTextSharp.text.Rectangle(llx, lly, urx, ury), dataGridView1.Rows[i].Cells[0].Value.ToString(), "Yes");
                        fCell.CheckType = RadioCheckField.TYPE_CROSS;
                        PdfFormField footerCheck = null;
                        footerCheck = fCell.CheckField;
                        stamp.AddAnnotation(footerCheck, x);
                    }
                    else
                    {
                    }

                }
                //Loop Ends here
                
                stamp.Close();
                fs.Close();
                reader.Close();
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Parse(string PDFInput)
        {

            try
            {
                dataGridView1.Rows.Clear();
                PdfFileAnalyzer.PdfReader Reader = new PdfFileAnalyzer.PdfReader();

                Reader.OpenPdfFile(PDFInput);

                


                int page_number = 0;

                for (int Index = Reader.ObjectArray.Length - 1; Index >= 0; Index--)
                {
                    ///get pages
                    if ((Reader.ObjectArray[Index] != null) && (Reader.ObjectArray[Index].PdfObjectType == "/Pages"))
                    {

                        PdfFileAnalyzer.PdfIndirectObject ReaderObject_pages = Reader.ObjectArray[Index];

                        var object_summary_pages = Reports.ObjectSummary(ReaderObject_pages).ToString().Replace(System.Environment.NewLine, "");
                        var object_summary_pages1 = Regex.Match(object_summary_pages, @"\[([^)]*)\]").Groups[1].Value;
                        var object_summary_finalpage = object_summary_pages1.Replace(" 0 R ", ",");
                        var page_id = object_summary_finalpage.Replace(" 0 R", ",");

                        var new_page_id = page_id.Split(new[] { ',' }, System.StringSplitOptions.RemoveEmptyEntries);


                        //int page_number = 0;

                        foreach (string page_items in new_page_id)
                        {

                            try
                            {
                                PdfFileAnalyzer.PdfIndirectObject ReaderObject = Reader.ObjectArray[Convert.ToInt32(page_items)];
                                var object_summary = Reports.ObjectSummary(ReaderObject).ToString().Replace(System.Environment.NewLine, "");
                                int pFrom = object_summary.IndexOf("Annots ") + "key : ".Length;
                                int pTo = object_summary.LastIndexOf(" 0 R/Contents");
                                var annot_object_summary = object_summary.Substring(pFrom, pTo - pFrom).Replace(" ", "");
                                page_number = page_number + 1;

                                //get annot
                                PdfFileAnalyzer.PdfIndirectObject ReaderObject_annot = Reader.ObjectArray[Convert.ToInt32(annot_object_summary)];

                                var object_summary_annot = Reports.ObjectSummary(ReaderObject_annot).ToString().Replace(System.Environment.NewLine, "");
                                var object_summary_annot1 = Regex.Match(object_summary_annot, @"\[([^)]*)\]").Groups[1].Value;
                                var object_summary_annots = object_summary_annot1.Replace(" 0 R ", ",");
                                var stream_id = object_summary_annots.Replace(" 0 R", ",");

                                var new_strim_id = stream_id.Split(new[] { ',' }, System.StringSplitOptions.RemoveEmptyEntries);

                                

                                foreach (string items in new_strim_id)
                                {
                                    //get stream
                                    PdfFileAnalyzer.PdfIndirectObject ReaderObject_stream = Reader.ObjectArray[Convert.ToInt32(items)];

                                    var object_summary_stream = Reports.ObjectSummary(ReaderObject_stream).ToString().Replace(System.Environment.NewLine, "");
                                    int pFrom_stream = object_summary_stream.IndexOf("sData ") + "key : ".Length;
                                    int pTo_stream = object_summary_stream.IndexOf(" 0 R");
                                    var view_stream_id = object_summary_stream.Substring(pFrom_stream, pTo_stream - pFrom_stream).Replace(" ", "");

                                    //get rect
                                    var pattern1 = @"\Annots(.*?)\ R";

                                    Regex rgxA = new Regex(pattern1);
                                    string replacementtextA1 = Reports.ObjectSummary(ReaderObject_stream).Replace(" 0", "");
                                    var matchA = rgxA.Match(replacementtextA1);
                                    var patternoutA = matchA.Groups[1].Value;
                                    var pattern2 = @"\[(.*?)\]";

                                    Regex rgxB = new Regex(pattern2);
                                    var matchB = rgxB.Match(Reports.ObjectSummary(ReaderObject_stream));
                                    var patternoutA2 = matchB.Groups[1].Value;
                                    //get fileds variable
                                    PdfFileAnalyzer.PdfIndirectObject ReaderObject_variable = Reader.ObjectArray[Convert.ToInt32(view_stream_id)];
                                    StreamByteArray = ReaderObject_variable.ReadStream();
                                    byte[] TempByteArray = ReaderObject_variable.DecompressStream(StreamByteArray);
                                    StreamByteArray = TempByteArray;

                                    

                                    
                                    if (Reports.ByteArrayToString(StreamByteArray).Contains("STARTTAG"))
                                        
                                    {
                                        //regex
                                        var pattern = @"\STARTTAG(.+)ENDTAG";
                                        //var pattern = @"\DEFAULT(.+CO)|\DEFAULT(.+TE)|\DEFAULT(.+MC)|\DEFAULT(.+TF)|\DEFAULT(.+DA)";
                                        Regex rgx = new Regex(pattern);
                                        string replacementtext = Regex.Replace(Reports.ByteArrayToString(StreamByteArray), @"\t|\n|\r", "");
                                        string replacementtext1 = replacementtext.Replace(".", "").Replace("!", "").Replace("%", "").Replace("(", "").Replace(")", "").Replace("'", "").Replace("&", "").Replace("#", "").Replace("$", "");
                                        var match = rgx.Match(replacementtext1);
                                        var patternout = match.Groups[1].Value;
                                        var patternout2 = patternout.Split(new char[] {':'})[0];

                                        //txtBox2.Text = txtBox2.Text + Convert.ToString(page_number) + " --> "+ patternout + patternout2 + patternout3 + patternout4 + patternout5 + "-->" + patternoutA2 + Environment.NewLine + Environment.NewLine;
                                        //MessageBox.Show(Convert.ToString(page_number) + patternout + patternout2 + patternout3 + patternout4 + patternout5 + "-->" + patternoutA2);
                                        DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();

                                        //row.Cells[0].Value = ValidateMe(patternout);
                                        row.Cells[0].Value = ValidateMe(patternout2.Trim());

                                        row.Cells[1].Value = patternoutA2;
                                        row.Cells[2].Value = Convert.ToString(page_number);

                                        string complocation = patternoutA2.ToString();

                                        float llx = float.Parse(complocation.Split(' ')[0]);
                                        float lly = float.Parse(complocation.Split(' ')[1]);
                                        float urx = float.Parse(complocation.Split(' ')[2]);
                                        float ury = float.Parse(complocation.Split(' ')[3]);

                                        if ((urx - llx) > 15)
                                        {
                                            row.Cells[3].Value = "TE";
                                        }
                                        else
                                        {
                                            row.Cells[3].Value = "CB";
                                        }

                                        dataGridView1.Rows.Add(row);
                                    }
                                    else
                                    {

                                        DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                                        row.Cells[0].Value = ToCamelCase(Reports.ByteArrayToString(StreamByteArray));
                                        row.Cells[1].Value = patternoutA2;
                                        row.Cells[2].Value = Convert.ToString(page_number);

                                        //string smartDetectCB = patternoutA2.Split(' ');

                                        string complocation = patternoutA2.ToString();

                                        float llx = float.Parse(complocation.Split(' ')[0]);
                                        float lly = float.Parse(complocation.Split(' ')[1]);
                                        float urx = float.Parse(complocation.Split(' ')[2]);
                                        float ury = float.Parse(complocation.Split(' ')[3]);

                                        if ((urx - llx) > 15)
                                        {
                                            row.Cells[3].Value = "TE";
                                        }
                                        else
                                        {
                                            row.Cells[3].Value = "CB";
                                        }


                                        dataGridView1.Rows.Add(row);

                                    }

                                }

                            }
                            catch
                            {

                            }


                        }

                    }





                }
               
                Reader.Dispose();
                CreateOutput(PDFInput);
                //End here
            }
            catch (Exception ex)
            {

                
                MessageBox.Show(ex.ToString());
                
            }
            
        }



        

        public string ValidateMe(String MyStr)
        {
            //MessageBox.Show(MyStr);
            string value = "";
            string MyStr2 = "";
            
            if (MyDict.TryGetValue(MyStr, out value))
            {
                
                MyStr2 = value;
            }
            else
            {
                MyStr2 = "*" + MyStr;
            }
            
            //validationEND------------------------------------------------------------------------------
            string ValidateMe = MyStr2;
            //MessageBox.Show(ValidateMe);
            return ValidateMe;
        }

        public static string ToCamelCase(String Str)
        {
            TextInfo txtInfo = new CultureInfo("en-us", false).TextInfo;
            Str = txtInfo.ToTitleCase(Str).Replace(" ", string.Empty);
            string ToCamelCase = Str;
            return ToCamelCase;
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            Init();
        }

        private void Init()
        {
            if (File.Exists(@ConfigPath))
            {
                dataSet = GetDataSetFromExcelFile(@ConfigPath);
                for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                {
                    if (!MyDict.ContainsKey(dataSet.Tables[0].Rows[i][1].ToString()))
                    {
                        MyDict.Add(dataSet.Tables[0].Rows[i][1].ToString(), dataSet.Tables[0].Rows[i][0].ToString());
                    }

                    //MessageBox.Show(dataSet.Tables[0].Rows[i][1].ToString() + "--" + dataSet.Tables[0].Rows[i][0].ToString());
                }
            }
            else
            {

                MessageBox.Show("Missing logic file in Config folder: " + Environment.NewLine + ConfigPath);
                this.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            String InputPath = AppDomain.CurrentDomain.BaseDirectory + "Input";
            String OutputPath = AppDomain.CurrentDomain.BaseDirectory + "Output";

            if (!Directory.Exists(InputPath))
            {
                Directory.CreateDirectory(InputPath);
            }
            if (!Directory.Exists(OutputPath))
            {
                Directory.CreateDirectory(OutputPath);
            }

            string[] InputPaths = Directory.GetFiles(InputPath);
            foreach (string filePath in InputPaths)
            {
                FileAttributes attrs = File.GetAttributes(filePath);
                if (attrs.HasFlag(FileAttributes.ReadOnly))
                    File.SetAttributes(filePath, attrs & ~FileAttributes.ReadOnly);
                File.Delete(filePath);
            }
            string[] OutputPaths = Directory.GetFiles(OutputPath);
            foreach (string filePath in OutputPaths)
            {
                FileAttributes attrs = File.GetAttributes(filePath);
                if (attrs.HasFlag(FileAttributes.ReadOnly))
                    File.SetAttributes(filePath, attrs & ~FileAttributes.ReadOnly);
                File.Delete(filePath);
            }

            openFileDialog2.Filter = "HPT files | *.hpt";
            openFileDialog2.Title = "Select HPT file";
            openFileDialog2.Multiselect = true;

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {

                foreach (string fileName in openFileDialog2.FileNames)
                {
                    try
                    {
                        File.Copy(fileName, Path.ChangeExtension(InputPath + @"\" + System.IO.Path.GetFileName(fileName), ".pdf"), true);

                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                    }

                }

                string[] FiletoProcessPaths = Directory.GetFiles(InputPath);
                foreach (string filePath in FiletoProcessPaths)
                {

                    Parse(filePath);

                }
            }
            if (MessageBox.Show("Conversion is complete." + Environment.NewLine + "Open file location?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(AppDomain.CurrentDomain.BaseDirectory + "Output");
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel files | *.xlsx";
            //openFileDialog1.Filter = "All files | *.*";
            openFileDialog1.Title = "Select Excel file";
            openFileDialog1.Multiselect = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    textBox1.Text = openFileDialog1.FileName;

                    //var dataSet = GetDataSetFromExcelFile(openFileDialog1.FileName);

                    //dataGridView2.AutoGenerateColumns = true;
                    //dataGridView2.DataSource = dataSet.Tables[0];
                    
                    //MessageBox.Show(string.Format("reading file: {0}", textBox1.Text));
                    //MessageBox.Show(string.Format("coloums: {0}", dataSet.Tables[0].Columns.Count));
                    //MessageBox.Show(string.Format("rows: {0}", dataSet.Tables[0].Rows.Count));
                    //Console.ReadKey();
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }
                
            }
        }
        private static string GetConnectionString(string file)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            string extension = file.Split('.').Last();

            if (extension == "xls")
            {
                //Excel 2003 and Older
                props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
                props["Extended Properties"] = "Excel 8.0";
            }
            else if (extension == "xlsx")
            {
                //Excel 2007, 2010, 2012, 2013
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
                props["Extended Properties"] = "Excel 12.0 XML";
            }
            else
                throw new Exception(string.Format("error file: {0}", file));

            props["Data Source"] = file;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        private static DataSet GetDataSetFromExcelFile(string file)
        {
            DataSet ds = new DataSet();

            string connectionString = GetConnectionString(file);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }

            return ds;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
    }

}



