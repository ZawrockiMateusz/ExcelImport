using GroupDocs.Conversion;
using GroupDocs.Conversion.FileTypes;
using GroupDocs.Conversion.Options.Convert;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

namespace ExcelImport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtExcelFilePath.Text = dlg.SelectedPath;
            }
            
        }
        private void btnBrowse2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtCsvFilePath.Text = dlg.SelectedPath;
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            var dataSource = txtDataSource.Text;
            var excelFilePath = txtExcelFilePath.Text;
            var databaseName = txtDatabaseName.Text;
            var tableName = txtTableName.Text;

            var csvFilePath = txtCsvFilePath.Text;

            using (Converter converter = new Converter(excelFilePath))
            {
                SpreadsheetConvertOptions options = new SpreadsheetConvertOptions
                {
                    PageNumber = 2,
                    PagesCount = 1,
                    Format = SpreadsheetFileType.Csv
                };
                converter.Convert(csvFilePath, options);



                var lineNumber = 0;
                using (SqlConnection conn = new SqlConnection(@"data source=" + dataSource + ";initial catalog=" + databaseName + ";trusted_connection=true")) //establishing connection string 
                {
                    conn.Open(); //opening connection with db
                    using (StreamReader reader = new StreamReader(@csvFilePath)) //path for .csv file
                    {
                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();

                            if (lineNumber != 0)
                            {
                                var values = line.Split(';');

                                var sql = "INSERT INTO " + databaseName + ".dbo." + tableName + " VALUES ('" + values[0] + "','" + values[1] + "'," + values[2] + ")"; //preparing sql query

                                //inserting .csv data into database
                                var cmd = new SqlCommand();
                                cmd.CommandText = sql;
                                cmd.CommandType = System.Data.CommandType.Text;
                                cmd.Connection = conn;
                                cmd.ExecuteNonQuery();
                            }
                            lineNumber += 1;
                        }
                    }
                    conn.Close(); //closing connection with db
                }
                MessageBox.Show("Data succesfully imported.");
            }
        }
    }
}
