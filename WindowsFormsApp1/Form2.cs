using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Windows.Forms.DataVisualization.Charting;
using System.Data.OleDb;
using MySqlX.XDevAPI.Relational;
using Org.BouncyCastle.Utilities.Collections;
using System.Diagnostics;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        

        public Form2()
        {
            InitializeComponent();

            // Assign the data to a field or property of the form
            
        }

        public void loadTabelV() 
        {
            DataTable dataTable = new DataTable();
            string query = "SELECT * FROM valuta";
            MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
            // Create a data adapter to fill the DataTable with the data
            MySqlDataAdapter adapter = new MySqlDataAdapter(command);
            adapter.Fill(dataTable);

            // Display the DataTable in a DataGridView control
            //DataGridView dataGridView = new DataGridView();
            dataGridView1.Columns.Clear();
            dataGridView1.DataSource = dataTable;
        }

        public void loadTabelB()
        {
            DataTable dataTable = new DataTable();
            string query = "SELECT * FROM beneficiar";
            MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
            // Create a data adapter to fill the DataTable with the data
            MySqlDataAdapter adapter = new MySqlDataAdapter(command);
            adapter.Fill(dataTable);

            // Display the DataTable in a DataGridView control
            //DataGridView dataGridView = new DataGridView();
            dataGridView2.Columns.Clear();
            dataGridView2.DataSource = dataTable;
        }
        public void loadTabelL()
        {
            DataTable dataTable = new DataTable();
            string query = "SELECT * FROM casadeschimb";
            MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
            // Create a data adapter to fill the DataTable with the data
            MySqlDataAdapter adapter = new MySqlDataAdapter(command);
            adapter.Fill(dataTable);

            // Display the DataTable in a DataGridView control
            //DataGridView dataGridView = new DataGridView();
            dataGridView3.Columns.Clear();
            dataGridView3.DataSource = dataTable;
        }

        public void loadTabelC()
        {
            DataTable dataTable = new DataTable();
            string query = "select idcom, user,numeutil,rol,comanda,dataCom,u.id from comenzi c,util u where c.id=u.id";
            MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
            // Create a data adapter to fill the DataTable with the data
            MySqlDataAdapter adapter = new MySqlDataAdapter(command);
            adapter.Fill(dataTable);

            // Display the DataTable in a DataGridView control
            //DataGridView dataGridView = new DataGridView();
            dataGridView4.Columns.Clear();
            dataGridView4.DataSource = dataTable;
        }

        public void loadCombo()
        {
            
            string query = "select distinct denumire from valuta";
            MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
           /* command.Parameters.AddWithValue("@valuta", valutaName);*/

            MySqlDataAdapter adapter = new MySqlDataAdapter(command);
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);

            // Clear the ComboBox before adding new items
            comboBox1.Items.Clear();

            // Loop through the rows in the DataTable and add the value of the "columnName" column to the ComboBox
            foreach (DataRow row in dataTable.Rows)
            {
                comboBox1.Items.Add(row["denumire"].ToString());
            }

        }

        public void loadChart()
        {
            MySqlConnection connection = ConnectionJDBC.GetConnection();
            string query = "select valutaschimb as Valuta, sum(suma) as Suma from beneficiar group by valutaschimb";
            MySqlCommand cmd = new MySqlCommand(query, connection);
            MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);

            // Create a new chart control and set the data source
            Chart chart1 = new Chart();
            chart1.DataSource = dataTable;

            chart1.Dock = DockStyle.Fill;
            chart1.BackColor = System.Drawing.Color.Transparent;

            // Specify the series and data points for the chart
            Series series = new Series("Valuta Schimb by Category");
            series.ChartType = SeriesChartType.Pie;
            series.XValueMember = "Valuta";
            series.YValueMembers = "Suma";
            series["PieLabelStyle"] = "Outside";
            series["PieStartAngle"] = "90";
            series["CollectedThreshold"] = "5";
            series["CollectedLabel"] = "#OTHER{P2}";
            series["CollectedLegendText"] = "Other";
            series.CustomProperties = "PieLabelStyle=Outside,LabelsRadialLineSize=1.2,LabelsHorizontalLineSize=1.3,LabelsVerticalLineSize=1.2";
            series.Label = "#VALX: #PERCENT{P2}";

            chart1.Series.Add(series);

            // Customize the appearance and layout of the chart
            chart1.Legends.Add(new Legend("Legend"));
            chart1.Titles.Add(new Title("Valuta Schimb by Category"));
            chart1.ChartAreas.Add(new ChartArea("ChartArea"));
            chart1.ChartAreas[0].Area3DStyle.Enable3D = true;
            chart1.ChartAreas[0].Area3DStyle.Inclination = 45;
            chart1.ChartAreas[0].Area3DStyle.Rotation = 45;
            chart1.DataBind();

            // Add the chart control to your form or user control
            tabControl1.TabPages[4].Controls.Add(chart1);

        }

        public void adaugCom(string data,string comanda)
        {
            try
            {
                DateTime currentDate = DateTime.Now;
                DataTable dataTable = new DataTable();
                string query = "insert into comenzi(user,comanda,dataCom) values(@user,@comanda,@dataCom)";
                MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
                /*command.Parameters.AddWithValue("@codv", );*/
                command.Parameters.AddWithValue("@user", data);
                command.Parameters.AddWithValue("@comanda", comanda);
                command.Parameters.AddWithValue("@ziua", currentDate);
                


                // Create a data adapter with the insert command
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(dataTable);

                // Display the DataTable in a DataGridView control
                //DataGridView dataGridView = new DataGridView();
                dataGridView4.Columns.Clear();
                dataGridView4.DataSource = dataTable;

                // Refresh the DataGridView to show the new row
                loadTabelC();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error inserting data: " + ex.Message);
            }
        }

        public void delTabelV()
        {
            
            if (dataGridView1.SelectedRows.Count > 0)
            {
                // Get the selected row
                DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];

                // Get the ID of the selected row from the appropriate column
                int id = Convert.ToInt32(selectedRow.Cells["codv"].Value); 
                DataTable dataTable = new DataTable();
                string query = "DELETE FROM valuta where codv=@codv";
                MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
                // Create a data adapter to fill the DataTable with the data
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                command.Parameters.AddWithValue("@codv", id);
                adapter.Fill(dataTable);
                // Remove the selected row from the DataGridView
                dataGridView1.Rows.Remove(selectedRow);
                dataGridView1.Columns.Clear();
                dataGridView1.DataSource = dataTable;
            }
            else
            {
                MessageBox.Show("Failed to delete the row from the database.");
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 m = new Form1();
            m.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            tabControl1.Visible = true;
            tabControl1.SelectedIndex = 0;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            tabControl1.Visible = true;
            tabControl1.SelectedIndex = 3;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            tabControl1.Visible = true;
            tabControl1.SelectedIndex = 1;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            tabControl1.Visible = true;
            tabControl1.SelectedIndex = 2;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            tabControl1.Visible = true;
            tabControl1.SelectedIndex =4;
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void Form2_Load(object sender, EventArgs e)
        {
            loadTabelV();
            loadTabelB();
            loadTabelL();
            loadTabelC();
            loadChart();
            loadCombo();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            /*// Check if a row is selected
            if (dataGridView1.SelectedRows.Count > 0)
            {
                // Get the selected row
                DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];

                // Remove the selected row from the DataGridView
                dataGridView1.Rows.Remove(selectedRow);
            }*/

            delTabelV(); 
            loadTabelV();
        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                int indexRand = dataGridView1.SelectedRows[0].Index;
                textBox1.Text = dataGridView1.Rows[indexRand].Cells[0].Value.ToString();
                textBox2.Text = dataGridView1.Rows[indexRand].Cells[1].Value.ToString();
                textBox3.Text = dataGridView1.Rows[indexRand].Cells[2].Value.ToString();
                textBox4.Text = dataGridView1.Rows[indexRand].Cells[3].Value.ToString();
                textBox5.Text = dataGridView1.Rows[indexRand].Cells[4].Value.ToString();
                textBox6.Text = dataGridView1.Rows[indexRand].Cells[5].Value.ToString();
                textBox7.Text = dataGridView1.Rows[indexRand].Cells[7].Value.ToString();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime currentDate = DateTime.Now;
                DataTable dataTable = new DataTable();
                string query = "INSERT INTO valuta(denumire,cursant,curscump,cursvanz,comision,ziua,codc) VALUES(@denumire,@cursant,@curscump,@cursvanz,@comision,@ziua,@codc)";
                MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
                /*command.Parameters.AddWithValue("@codv", );*/
                command.Parameters.AddWithValue("@denumire",textBox2.Text);
                command.Parameters.AddWithValue("@cursant", textBox3.Text);
                command.Parameters.AddWithValue("@curscump", textBox4.Text);
                command.Parameters.AddWithValue("@cursvanz", textBox5.Text);
                command.Parameters.AddWithValue("@comision", textBox6.Text);
                command.Parameters.AddWithValue("@ziua", currentDate);
                command.Parameters.AddWithValue("@codc", textBox7.Text);

                // Create a data adapter with the insert command
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(dataTable);

                // Display the DataTable in a DataGridView control
                //DataGridView dataGridView = new DataGridView();
                dataGridView1.Columns.Clear();
                dataGridView1.DataSource = dataTable;
               
                // Refresh the DataGridView to show the new row
                loadTabelV();
                string comanda = "Adaug valuta";
                /*adaugCom(data, comanda);*/
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error inserting data: " + ex.Message);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime currentDate = DateTime.Now;
                DataTable dataTable = new DataTable();
                string query = "UPDATE valuta set denumire=@denumire,cursant=@cursant,curscump=@curscump,cursvanz=@cursvanz,comision=@comision,codc=@codc where codv=@codv";
                MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
               /* command.Parameters.AddWithValue("@codv", textBox1.Text);*/
                command.Parameters.AddWithValue("@denumire", textBox2.Text);
                command.Parameters.AddWithValue("@cursant", textBox3.Text);
                command.Parameters.AddWithValue("@curscump", textBox4.Text);
                command.Parameters.AddWithValue("@cursvanz", textBox5.Text);
                command.Parameters.AddWithValue("@comision", textBox6.Text);
                command.Parameters.AddWithValue("@ziua", currentDate);
                command.Parameters.AddWithValue("@codc", textBox7.Text);
                command.Parameters.AddWithValue("@codv", textBox1.Text); // specify the primary key for the WHERE clause

                // Create a data adapter with the insert command
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(dataTable);

                // Display the DataTable in a DataGridView control
                //DataGridView dataGridView = new DataGridView();
                dataGridView1.Columns.Clear();
                dataGridView1.DataSource = dataTable;

                // Refresh the DataGridView to show the new row
                loadTabelV();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error inserting data: " + ex.Message);
            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count > 0)
            {
                int indexRand = dataGridView2.SelectedRows[0].Index;
                textBox16.Text = dataGridView2.Rows[indexRand].Cells[0].Value.ToString();
                textBox15.Text = dataGridView2.Rows[indexRand].Cells[1].Value.ToString();
                textBox14.Text = dataGridView2.Rows[indexRand].Cells[2].Value.ToString();
                textBox13.Text = dataGridView2.Rows[indexRand].Cells[3].Value.ToString();
                textBox12.Text = dataGridView2.Rows[indexRand].Cells[4].Value.ToString();
                textBox11.Text = dataGridView2.Rows[indexRand].Cells[5].Value.ToString();
                textBox10.Text = dataGridView2.Rows[indexRand].Cells[6].Value.ToString();
                textBox25.Text = dataGridView2.Rows[indexRand].Cells[8].Value.ToString();
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime currentDate = DateTime.Now;
                DataTable dataTable = new DataTable();
                string query = "INSERT INTO beneficiar(nume,pren,suma,valutaschimb,valutaprim,comisB,data,codc) VALUES(@nume, @pren, @suma, @valutaschimb, @valutaprim, @comisB, @data, @codc)";
                MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
                command.Parameters.AddWithValue("@nume", textBox15.Text);
                command.Parameters.AddWithValue("@pren", textBox14.Text);
                command.Parameters.AddWithValue("@suma", textBox13.Text);
                command.Parameters.AddWithValue("@valutaschimb", textBox12.Text);
                command.Parameters.AddWithValue("@valutaprim", textBox11.Text);
                command.Parameters.AddWithValue("@comisB", textBox10.Text);
                command.Parameters.AddWithValue("@data", currentDate);
                command.Parameters.AddWithValue("@codc", textBox25.Text);

                // Create a data adapter with the insert command
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(dataTable);

                // Display the DataTable in a DataGridView control
                //DataGridView dataGridView = new DataGridView();
                dataGridView1.Columns.Clear();
                dataGridView1.DataSource = dataTable;

                // Refresh the DataGridView to show the new row
                loadTabelB();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error inserting data: " + ex.Message);
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count > 0)
            {
                // Get the selected row
                DataGridViewRow selectedRow = dataGridView2.SelectedRows[0];

                // Get the ID of the selected row from the appropriate column
                int id = Convert.ToInt32(selectedRow.Cells["codbe"].Value);
                DataTable dataTable = new DataTable();
                string query = "DELETE FROM beneficiar where codbe=@codbe";
                MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
                // Create a data adapter to fill the DataTable with the data
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                command.Parameters.AddWithValue("@codbe", id);
                adapter.Fill(dataTable);
                // Remove the selected row from the DataGridView
                dataGridView2.Rows.Remove(selectedRow);
                dataGridView2.Columns.Clear();
                dataGridView2.DataSource = dataTable;
            }
            else
            {
                MessageBox.Show("Failed to delete the row from the database.");
            }
            loadTabelB() ;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime currentDate = DateTime.Now;
                DataTable dataTable = new DataTable();
                string query = "UPDATE beneficiar SET nume=@nume, pren=@pren, suma=@suma, valutaschimb=@valutaschimb, valutaprim=@valutaprim, comisB=@comisB, data=@data, codc=@codc WHERE codbe=@codbe";
                MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
                /* command.Parameters.AddWithValue("@codv", textBox1.Text);*/
                command.Parameters.AddWithValue("@nume", textBox15.Text);
                command.Parameters.AddWithValue("@pren", textBox14.Text);
                command.Parameters.AddWithValue("@suma", textBox13.Text);
                command.Parameters.AddWithValue("@valutaschimb", textBox12.Text);
                command.Parameters.AddWithValue("@valutaprim", textBox11.Text);
                command.Parameters.AddWithValue("@comisB", textBox10.Text);
                command.Parameters.AddWithValue("@data", currentDate);
                command.Parameters.AddWithValue("@codc", textBox25.Text);
                command.Parameters.AddWithValue("@codbe", textBox16.Text); // specify the primary key for the WHERE clause

                // Create a data adapter with the insert command
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(dataTable);

                // Display the DataTable in a DataGridView control
                //DataGridView dataGridView = new DataGridView();
                dataGridView2.Columns.Clear();
                dataGridView2.DataSource = dataTable;

                // Refresh the DataGridView to show the new row
                loadTabelB();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error inserting data: " + ex.Message);
            }
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.SelectedRows.Count > 0)
            {
                int indexRand = dataGridView3.SelectedRows[0].Index;
                textBox24.Text = dataGridView3.Rows[indexRand].Cells[0].Value.ToString();
                textBox23.Text = dataGridView3.Rows[indexRand].Cells[1].Value.ToString();
                textBox22.Text = dataGridView3.Rows[indexRand].Cells[2].Value.ToString();
                textBox21.Text = dataGridView3.Rows[indexRand].Cells[3].Value.ToString();
                textBox20.Text = dataGridView3.Rows[indexRand].Cells[4].Value.ToString();
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime currentDate = DateTime.Now;
                DataTable dataTable = new DataTable();
                string query = "INSERT INTO casadeschimb(nume,strada,oras,judet) VALUES(@nume,@strada,@oras,@judet)";
                MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
                command.Parameters.AddWithValue("@nume", textBox23.Text);
                command.Parameters.AddWithValue("@strada", textBox22.Text);
                command.Parameters.AddWithValue("@oras", textBox21.Text);
                command.Parameters.AddWithValue("@judet", textBox20.Text);

                // Create a data adapter with the insert command
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(dataTable);

                // Display the DataTable in a DataGridView control
                //DataGridView dataGridView = new DataGridView();
                dataGridView3.Columns.Clear();
                dataGridView3.DataSource = dataTable;

                // Refresh the DataGridView to show the new row
                loadTabelL();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error inserting data: " + ex.Message);
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            // Create a PDF document
            Document doc = new Document(PageSize.LETTER, 10f, 10f, 10f, 0f);
            PdfWriter.GetInstance(doc, new FileStream(@"C:\Users\Petru\Desktop\DataGridView3.pdf", FileMode.Create));

            // Open the document
            doc.Open();

            // Create a PDF table
            PdfPTable pdfTable = new PdfPTable(dataGridView3.ColumnCount);
            pdfTable.DefaultCell.Padding = 3;
            pdfTable.WidthPercentage = 100;
            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;

            // Add header rows
            foreach (DataGridViewColumn column in dataGridView3.Columns)
            {
                PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                pdfTable.AddCell(cell);
            }

            // Add data rows
            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    pdfTable.AddCell(cell.Value.ToString());
                }
            }

            // Add PDF table to document
            doc.Add(pdfTable);

            // Close the document
            doc.Close();

            MessageBox.Show("PDF file saved to DataGridView.pdf");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Files (.xlsx)|.xlsx";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;
                    filePath = filePath.EndsWith(".xlsx") ? filePath : filePath + ".xlsx";

                    using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                    {
                        IWorkbook wb = new XSSFWorkbook();
                        ISheet sheet = wb.CreateSheet("Sheet1");

                        IRow rowCol = sheet.CreateRow(0);
                        for (int i = 0; i < dataGridView1.ColumnCount; i++)
                        {
                            ICell cell = rowCol.CreateCell(i);
                            cell.SetCellValue(dataGridView1.Columns[i].HeaderText);
                        }

                        for (int j = 0; j < dataGridView1.RowCount; j++)
                        {
                            IRow row = sheet.CreateRow(j + 1);
                            for (int k = 0; k < dataGridView1.ColumnCount; k++)
                            {
                                ICell cell = row.CreateCell(k);
                                if (dataGridView1.Rows[j].Cells[k].Value != null)
                                {
                                    cell.SetCellValue(dataGridView1.Rows[j].Cells[k].Value.ToString());
                                }
                            }
                        }

                        wb.Write(fs);
                        wb.Close();
                        fs.Close();
                    }

                    Process.Start(filePath);
                }
                else
                {
                    MessageBox.Show("Eroare salvare fisier excel");
                }
               
            }
            catch (IOException io)
            {
                Console.WriteLine(io.Message);
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            // Create a PDF document
            Document doc = new Document(PageSize.LETTER, 10f, 10f, 10f, 0f);
            PdfWriter.GetInstance(doc, new FileStream(@"C:\Users\Petru\Desktop\DataGridView2.pdf", FileMode.Create));

            // Open the document
            doc.Open();

            // Create a PDF table
            PdfPTable pdfTable = new PdfPTable(dataGridView2.ColumnCount);
            pdfTable.DefaultCell.Padding = 3;
            pdfTable.WidthPercentage = 100;
            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;

            // Add header rows
            foreach (DataGridViewColumn column in dataGridView2.Columns)
            {
                PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                pdfTable.AddCell(cell);
            }

            // Add data rows
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    pdfTable.AddCell(cell.Value.ToString());
                }
            }

            // Add PDF table to document
            doc.Add(pdfTable);

            // Close the document
            doc.Close();

            MessageBox.Show("PDF file saved to DataGridView.pdf");
        }

        private void button15_Click(object sender, EventArgs e)
        {
            // Create a new Excel file
            var fileName = "DataGridView2.xlsx";
            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);
            var spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);

            // Add a new worksheet
            var workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
            sheets.Append(sheet);

            // Add header row
            var headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
            foreach (DataGridViewColumn column in dataGridView2.Columns)
            {
                headerRow.AppendChild(new Cell(new InlineString(new Text(column.HeaderText))));
            }
            sheetData.AppendChild(headerRow);

            // Add data rows
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                var dataRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dataRow.AppendChild(new Cell(new InlineString(new Text(cell.Value.ToString()))));
                }
                sheetData.AppendChild(dataRow);
            }

            // Save the Excel file
            workbookPart.Workbook.Save();
            spreadsheetDocument.Close();

            MessageBox.Show("Excel file saved to " + filePath);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            // Create a PDF document
            Document doc = new Document(PageSize.LETTER, 10f, 10f, 10f, 0f);
            PdfWriter.GetInstance(doc, new FileStream(@"C:\Users\Petru\Desktop\DataGridView1.pdf", FileMode.Create));

            // Open the document
            doc.Open();

            // Create a PDF table
            PdfPTable pdfTable = new PdfPTable(dataGridView1.ColumnCount);
            pdfTable.DefaultCell.Padding = 3;
            pdfTable.WidthPercentage = 100;
            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;

            // Add header rows
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                pdfTable.AddCell(cell);
            }

            // Add data rows
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    pdfTable.AddCell(cell.Value.ToString());
                }
            }

            // Add PDF table to document
            doc.Add(pdfTable);

            // Close the document
            doc.Close();

            MessageBox.Show("PDF file saved to DataGridView.pdf");
        }

        private void button21_Click(object sender, EventArgs e)
        {
            // Create a new Excel file
            var fileName = "DataGridView3.xlsx";
            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);
            var spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);

            // Add a new worksheet
            var workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
            sheets.Append(sheet);

            // Add header row
            var headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
            foreach (DataGridViewColumn column in dataGridView3.Columns)
            {
                headerRow.AppendChild(new Cell(new InlineString(new Text(column.HeaderText))));
            }
            sheetData.AppendChild(headerRow);

            // Add data rows
            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                var dataRow = new   DocumentFormat.OpenXml.Spreadsheet.Row();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dataRow.AppendChild(new Cell(new InlineString(new Text(cell.Value.ToString()))));
                }
                sheetData.AppendChild(dataRow);
            }

            // Save the Excel file
            workbookPart.Workbook.Save();
            spreadsheetDocument.Close();

            MessageBox.Show("Excel file saved to " + filePath);
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox8.Text)) 
            {
                loadTabelV();
            }
            else
            {
                DataTable dataTable = new DataTable();
                string query = "SELECT * FROM valuta where denumire=@denumire";
                MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
                command.Parameters.AddWithValue("@denumire", textBox8.Text);
                // Create a data adapter to fill the DataTable with the data
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(dataTable);

                // Display the DataTable in a DataGridView control
                //DataGridView dataGridView = new DataGridView();
                dataGridView1.Columns.Clear();
                dataGridView1.DataSource = dataTable;
            }
            
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox9.Text))
            {
                if (comboBox1.SelectedIndex == -1)
                {
                    loadTabelB();
                }
                else
                {
                    DataTable dataTable = new DataTable();
                    string query = "SELECT * FROM beneficiar where valutaschimb=@valutaschimb";
                    MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
                    command.Parameters.AddWithValue("@valutaschimb", comboBox1.SelectedItem);
                    // Create a data adapter to fill the DataTable with the data
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(dataTable);

                    // Display the DataTable in a DataGridView control
                    //DataGridView dataGridView = new DataGridView();
                    dataGridView2.Columns.Clear();
                    dataGridView2.DataSource = dataTable;
                }
            }
            else
            {
                if (comboBox1.SelectedIndex == -1)
                {
                    DataTable dataTable = new DataTable();
                    string query = "SELECT * FROM beneficiar where nume=@nume";
                    MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
                    command.Parameters.AddWithValue("@nume", textBox9.Text);
                    // Create a data adapter to fill the DataTable with the data
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(dataTable);

                    // Display the DataTable in a DataGridView control
                    //DataGridView dataGridView = new DataGridView();
                    dataGridView2.Columns.Clear();
                    dataGridView2.DataSource = dataTable;
                }
                else 
                {
                    DataTable dataTable = new DataTable();
                    string query = "SELECT * FROM beneficiar where valutaschimb=@valutaschimb and nume=@nume";
                    MySqlCommand command = new MySqlCommand(query, ConnectionJDBC.GetConnection());
                    command.Parameters.AddWithValue("@valutaschimb", comboBox1.SelectedItem);
                    command.Parameters.AddWithValue("@nume", textBox9.Text);
                    // Create a data adapter to fill the DataTable with the data
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(dataTable);

                    // Display the DataTable in a DataGridView control
                    //DataGridView dataGridView = new DataGridView();
                    dataGridView2.Columns.Clear();
                    dataGridView2.DataSource = dataTable;
                }
            }
            
        }

        private void button20_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (.xlsx)|.xlsx|All files (.)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFileDialog.FileName;
                string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;'";
                OleDbConnection connection = new OleDbConnection(connectionString);
                OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connection);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }
        }

        }

        


    
}
