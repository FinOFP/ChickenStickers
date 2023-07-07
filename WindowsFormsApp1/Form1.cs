using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Seagull.BarTender.Print;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        // Dictionary to hold SKU as key and Product Data as value
        private Dictionary<string, (string ProductName, double WeeklyAverage, double Monday, double Tuesday, double Wednesday, double Thursday, double Friday)> _data;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Initialize product data
            InitializeData();

            // Populate the DataGridView with the product data
            PopulateDataGridView();

            // Attach CellValueChanged event handler
            dataGridView1.CellValueChanged += DataGridView1_CellValueChanged;

            // Set the visual styles
            dataGridView1.RowsDefaultCellStyle.BackColor = Color.FromArgb(235, 235, 235);
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font(dataGridView1.Font, FontStyle.Bold);

            // Hide the row headers
            dataGridView1.RowHeadersVisible = false;

            // Set font
            dataGridView1.DefaultCellStyle.Font = new Font("Segoe UI", 10);

            dataGridView1.CellFormatting += DataGridView1_CellFormatting;

        }


        private void InitializeData()
        {
            // Populate the product data
            // Tuple structure: (ProductName, WeeklyAverage, Monday, Tuesday, Wednesday, Thursday, Friday)
            _data = new Dictionary<string, (string, double, double, double, double, double, double)>
            {
                { "118s1", ("Plain Sliced 1kg", 106.0, 10.0, 22.0, 35.0, 20.0, 21.0) },
                { "118s2", ("Plain Sliced 2kg", 41.0, 15.0, 3.0, 7.0, 15.0, 3.0) },
                { "118s8", ("Plain Sliced 8kg", 7.0, 0.0, 0.0, 0.0, 7.0, 0.0) },
                { "118th", ("Thin Plain Sliced 2kg", 38.0, 0.0, 18.0, 0.0, 9.0, 11.0) },
                { "118TA1", ("Tandoori Breast Sliced 1kg", 20.0, 2.0, 2.0, 8.0, 1.0, 8.0) },
                { "118LE", ("Lemon & Herb Breast Sliced 1kg", 22.0, 6.0, 5.0, 5.0, 3.0, 5.0) },
                { "118po", ("Portugese Breast Sliced 1kg", 0.0, 0.0, 0.0, 0.0, 0.0, 0.0) },
                { "118PE", ("Pesto Breast Sliced 1kg", 9.0, 1.0, 3.0, 3.0, 1.0, 2.0) },
                { "118PM", ("Pesto Mayo Breast Sliced 1kg", 24.0, 4.0, 8.0, 4.0, 4.0, 6.0) },
                { "cmm", ("Chicken & Mayo Sliced 1kg", 16.0, 2.0, 4.0, 3.0, 3.0, 5.0) },
                { "fwcm", ("Chicken & Mayo Sliced 2kg", 59.0, 8.0, 14.0, 15.0, 16.0, 7.0) },
                { "118pp", ("Peri Mayo Breast Sliced 1kg", 19.0, 4.0, 8.0, 4.0, 1.0, 3.0) },
                { "fw118pp", ("Peri (no mayo) Breast Sliced 2kg", 26.0, 3.0, 12.0, 7.0, 3.0, 2.0) },
                { "fwcc", ("Coronation Breast Sliced 2kg", 7.0, 1.0, 3.0, 2.0, 2.0, 1.0) },
                { "cmwt", ("Walnut Tarragon Breast Sliced 2kg", 6.0, 0.0, 5.0, 0.0, 1.0, 0.0) },
                { "118TE", ("Teriyaki Breast Sliced 1kg", 5.0, 0.0, 2.0, 0.0, 1.0, 2.0) },
                { "tuma1", ("Tuna & Mayo 1kg", 31.0, 5.0, 11.0, 3.0, 7.0, 6.0) },
                { "aioli", ("Aioli Bag 2kg", 4.0, 1.0, 1.0, 1.0, 0.0, 2.0) },
                { "tenp1", ("Plain Tenderloins 1kg", 5.0, 0.0, 2.0, 0.0, 2.0, 2.0) },
                { "118mex", ("Achiote Mex Tenderloin 1kg", 15.0, 3.0, 4.0, 1.0, 5.0, 3.0) },
                { "1031", ("Poached Breast 500g", 124.0, 19.0, 15.0, 30.0, 4.0, 58.0) },
                { "1221", ("Crumbed Tenderloins 1kg", 23.0, 3.0, 6.0, 3.0, 6.0, 6.0) },
                { "119cr", ("Crumbed Schnitzel x6", 2.0, 1.0, 2.0, 0.0, 0.0, 0.0) },
                { "119pl", ("Plain Schnitzel x6", 1.0, 0.0, 0.0, 0.0, 0.0, 1.0) },
                { "119LE", ("Lemon & Herb Schnitzel x6", 13.0, 1.0, 2.0, 5.0, 3.0, 2.0) },
                { "119pp", ("Peri Peri Schnitzel x6", 16.0, 1.0, 1.0, 2.0, 4.0, 8.0) },
                { "119TA", ("Tandoori Schnitzel x6", 0.0, 0.0, 0.0, 0.0, 0.0, 0.0) },
                { "stgb", ("St. George Box (8x1kg)", 0.0, 0.0, 0.0, 0.0, 0.0, 0.0) }
            };
        }

        private void PopulateDataGridView()
        {
            dataGridView1.RowTemplate.Height = 23;
            // Add columns to the DataGridView
            dataGridView1.Columns.Add("sku", "SKU");
            dataGridView1.Columns.Add("productName", "Product Name");
            dataGridView1.Columns.Add("weeklyAverage", "Weekly Average");
            dataGridView1.Columns.Add("tomorrowAverage", "Tomorrow Average");
            dataGridView1.Columns.Add("stockLevel", "Stock Level");
            dataGridView1.Columns.Add("need", "Need");
            dataGridView1.Columns.Add("stickers", "Stickers");

            //setting column widths
            dataGridView1.Columns["sku"].Width = 65; // 60 pixels wide for the 'SKU' column
            dataGridView1.Columns["productName"].Width = 250; // auto adjsut this one
            dataGridView1.Columns["weeklyAverage"].Width = 65; // 70 pixels wide for the 'Weekly Average' column
            dataGridView1.Columns["tomorrowAverage"].Width = 65; // 70 pixels wide for the 'Weekly Average' column
            dataGridView1.Columns["stockLevel"].Width = 65; // 70 pixels wide for the 'Weekly Average' column
            dataGridView1.Columns["need"].Width = 65; // 70 pixels wide for the 'Weekly Average' column
            dataGridView1.Columns["stickers"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; // auto adjsut this one

            // Set ReadOnly properties for certain columns
            dataGridView1.Columns["stockLevel"].ReadOnly = false;
            dataGridView1.Columns["stickers"].ReadOnly = false;

            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                if (column.Name != "stockLevel" && column.Name != "stickers")
                {
                    column.ReadOnly = true;
                }
            }

            // Determine the next weekday
            DateTime nextWeekday = GetNextWeekday(DateTime.Now);

            // Add rows to the DataGridView
            foreach (var item in _data)
            {
                double tomorrowAverage = 0;

                // Set the value of tomorrowAverage based on the next weekday
                switch (nextWeekday.DayOfWeek)
                {
                    case DayOfWeek.Monday:
                        tomorrowAverage = item.Value.Monday;
                        break;
                    case DayOfWeek.Tuesday:
                        tomorrowAverage = item.Value.Tuesday;
                        break;
                    case DayOfWeek.Wednesday:
                        tomorrowAverage = item.Value.Wednesday;
                        break;
                    case DayOfWeek.Thursday:
                        tomorrowAverage = item.Value.Thursday;
                        break;
                    case DayOfWeek.Friday:
                        tomorrowAverage = item.Value.Friday;
                        break;
                }

                // Add a new row to the DataGridView
                int rowIndex = dataGridView1.Rows.Add(item.Key, item.Value.ProductName, item.Value.WeeklyAverage, tomorrowAverage, "", "", "");

                // Set default values for editable cells
                var stockLevelCell = dataGridView1.Rows[rowIndex].Cells["stockLevel"];
                var stickersCell = dataGridView1.Rows[rowIndex].Cells["stickers"];

                stockLevelCell.Value = "";
                stickersCell.Value = "";
            }
        }


        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                StringBuilder summary = new StringBuilder(); // Keep track of printed stickers
                Print.Text = "Printing In Progress...";
                Print.Enabled = false;

                using (Seagull.BarTender.Print.Engine engine = new Seagull.BarTender.Print.Engine(true))
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        var skuCell = row.Cells["sku"];
                        var stickersCell = row.Cells["stickers"];
                        var productCell = row.Cells["productName"];

                        if (skuCell.Value != null)
                        {
                            string sku = skuCell.Value.ToString();

                            // Check if the stickers cell is not blank and contains a valid number
                            if (stickersCell.Value == null ||
                                string.IsNullOrWhiteSpace(stickersCell.Value.ToString()) ||
                                !int.TryParse(stickersCell.Value.ToString(), out int numCopies) ||
                                numCopies < 1)
                            {
                                continue; // Skip to the next row
                            }

                            if (sku == "stgb")
                            {
                                // Print "stgb" SKU
                                PrintLabel(engine, sku, numCopies * 2); // 2 times the sticker value
                                summary.AppendLine($"{sku}, {productCell.Value} was printed {numCopies * 2} times."); // Append to the summary

                                // Print "stgb1" SKU
                                PrintLabel(engine, "stgb1", numCopies * 8); // 8 times the sticker value
                                summary.AppendLine($"stgb1, {productCell.Value} was printed {numCopies * 8} times."); // Append to the summary
                            }
                            else
                            {
                                // Print normal SKU
                                PrintLabel(engine, sku, numCopies);
                                summary.AppendLine($"{sku}, {productCell.Value} was printed {numCopies} times."); // Append to the summary
                            }
                        }
                    }
                }

                // Show the summary in a message box after printing
                if (summary.Length > 0)
                {
                    MessageBox.Show(summary.ToString());
                }
                Print.Text = "Print";
                Print.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            // Initialize a PrintDocument for printing the table
            var printDocument = new PrintDocument();
            printDocument.PrintPage += PrintDocument_PrintPage;

            // Show the Print Dialog and Print the table
            var printDialog = new PrintDialog();
            printDialog.Document = printDocument;
            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                printDocument.Print();
            }
            Print.Text = "Print";
            Print.Enabled = true;
        }

        private Dictionary<string, string> skuCategories = new Dictionary<string, string>
        {
            {"118s1", "Plain Sliced"},
            {"118s2", "Plain Sliced"},
            {"118s8", "Plain Sliced"},
            {"118th", "Plain Sliced"},
            {"118TA1", "Flavoured Sliced"},
            {"118LE", "Flavoured Sliced"},
            {"118po", "Flavoured Sliced"},
            {"118PE", "Flavoured Sliced"},
            {"118PM", "Flavoured Sliced"},
            {"cmm", "Flavoured Sliced"},
            {"fwcm", "Flavoured Sliced"},
            {"118pp", "Flavoured Sliced"},
            {"fw118pp", "Flavoured Sliced"},
            {"fwcc", "Flavoured Sliced"},
            {"cmwt", "Flavoured Sliced"},
            {"118TE", "Flavoured Sliced"},
            {"tuma1", "Tuna"},
            {"aioli", "Aioli"},
            {"tenp1", "Tenderloins & Breast"},
            {"118mex", "Tenderloins & Breast"},
            {"1031", "Tenderloins & Breast"},
            {"1221", "Crumbed"},
            {"119cr", "Crumbed"},
            {"119pl", "Schnitzel"},
            {"119LE", "Schnitzel"},
            {"119pp", "Schnitzel"},
            {"119TA", "Schnitzel"},
            {"stgb", "St. George"}
        };

        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            Graphics graphics = e.Graphics;

            // Set the font, its size, and cell sizes
            Font headerFont = new Font("Arial", 14, FontStyle.Bold);
            Font font = new Font("Arial", 12);
            int cellHeight = 30;
            int cellWidthProduct = 280;
            int cellWidthAmount = 80;

            // Define the starting position
            int startX = 10;
            int startY = 10;
            int offset = 0;

            // Add "Product" and "Amount" headers -- NEW
            graphics.DrawString("Product", headerFont, Brushes.Black, startX, startY + offset);
            graphics.DrawString("Amount", headerFont, Brushes.Black, startX + cellWidthProduct, startY + offset);
            offset += cellHeight;

            string lastCategory = null;

            // Loop through the DataGridView and print the data
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                DataGridViewRow row = dataGridView1.Rows[i];

                string sku = row.Cells["sku"].Value?.ToString() ?? string.Empty; // Change "productName" to "sku"
                string amount = row.Cells["stickers"].Value?.ToString();
                string productName = row.Cells["productName"].Value?.ToString() ?? string.Empty;

                // Replacing empty sticker fields with 0 -- NEW
                if (string.IsNullOrWhiteSpace(amount))
                {
                    amount = "0";
                }

                // Category header
                if (skuCategories.ContainsKey(sku) && skuCategories[sku] != lastCategory)
                {
                    lastCategory = skuCategories[sku];
                    graphics.DrawString(lastCategory, headerFont, Brushes.Black, startX, startY + offset);
                    offset += cellHeight;
                }

                // Cell background color
                int stickerCount = int.TryParse(amount, out stickerCount) ? stickerCount : 0;
                Color backgroundColor = Color.White;
                if (stickerCount > 1)
                {
                    backgroundColor = Color.PaleGoldenrod; // Change background color to PaleGoldenrod
                }
                else if (i % 2 == 0)
                {
                    backgroundColor = Color.LightGray;
                }
                graphics.FillRectangle(new SolidBrush(backgroundColor), startX, startY + offset, cellWidthProduct + cellWidthAmount, cellHeight);

                // Drawing the borders
                graphics.DrawRectangle(Pens.Black, startX, startY + offset, cellWidthProduct, cellHeight);
                graphics.DrawRectangle(Pens.Black, startX + cellWidthProduct, startY + offset, cellWidthAmount, cellHeight);

                // Drawing the content text
                StringFormat stringFormat = new StringFormat();
                stringFormat.Alignment = StringAlignment.Center;
                stringFormat.LineAlignment = StringAlignment.Center;

                RectangleF productCell = new RectangleF(startX, startY + offset, cellWidthProduct, cellHeight);
                RectangleF amountCell = new RectangleF(startX + cellWidthProduct, startY + offset, cellWidthAmount, cellHeight);

                graphics.DrawString(productName, font, Brushes.Black, productCell, stringFormat); // Change sku to productName
                graphics.DrawString(amount, font, Brushes.Black, amountCell, stringFormat);

                offset += cellHeight;
            }
        }

        private void PrintLabel(Seagull.BarTender.Print.Engine engine, string sku, int numCopies)
        {
            string filePath = @"C:\Users\finla\OneDrive\Desktop\Labels\" + sku + ".btw";
            Seagull.BarTender.Print.LabelFormatDocument format = engine.Documents.Open(filePath) as Seagull.BarTender.Print.LabelFormatDocument;

            if (format != null)
            {
                format.PrintSetup.IdenticalCopiesOfLabel = numCopies;
                Seagull.BarTender.Print.Messages messages;
                string printJobName = "My C# App Print Job";
                Seagull.BarTender.Print.Result result = format.Print(printJobName, out messages);
            }
        }

        private void DataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0) return; // Ignore changes to the header row

                // Handle changes in "stockLevel" column
                if (e.ColumnIndex == dataGridView1.Columns["stockLevel"].Index)
                {
                    var stockLevelCell = dataGridView1.Rows[e.RowIndex].Cells["stockLevel"];
                    var tomorrowAverageCell = dataGridView1.Rows[e.RowIndex].Cells["tomorrowAverage"];
                    var recommendedCell = dataGridView1.Rows[e.RowIndex].Cells["need"];

                    if (stockLevelCell.Value != null && tomorrowAverageCell.Value != null)
                    {
                        if (!int.TryParse(stockLevelCell.Value.ToString(), out int stockLevel))
                        {
                            MessageBox.Show("Please enter a valid number for Stock Level.");
                            stockLevelCell.Value = 0;
                            return;
                        }

                        double tomorrowAverage = Convert.ToDouble(tomorrowAverageCell.Value);
                        recommendedCell.Value = tomorrowAverage - stockLevel;
                    }
                }

                // Handle changes in "stickers" column
                if (e.ColumnIndex == dataGridView1.Columns["stickers"].Index)
                {
                    var cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];

                    if (cell.Value == null || string.IsNullOrWhiteSpace(cell.Value.ToString()))
                    {
                        cell.Value = 0;
                    }
                    else
                    {
                        if (!double.TryParse(cell.Value.ToString(), out _))
                        {
                            MessageBox.Show("Please enter a valid number for Stickers.");
                            cell.Value = 0;
                        }
                    }
                }
                Print.Text = "Print";
                Print.Enabled = true;
            }
            catch (Exception ex)
            {
                // Display the exception message.
                MessageBox.Show($"An error occurred: {ex.Message}");
                Print.Text = "Print";
                Print.Enabled = true;
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Your existing logic here...
        }
        private DateTime GetNextWeekday(DateTime date)
        {
            DateTime nextDay = date.AddDays(1);
            while (nextDay.DayOfWeek == DayOfWeek.Saturday || nextDay.DayOfWeek == DayOfWeek.Sunday)
            {
                nextDay = nextDay.AddDays(1);
            }
            return nextDay;
        }
        private void DataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return; // Ignore header row

            // Format the "need" column.
            if (dataGridView1.Columns[e.ColumnIndex].Name == "need")
            {
                if (e.Value != null)
                {
                    double value;
                    if (double.TryParse(e.Value.ToString(), out value))
                    {
                        if (value < -2)
                        {
                            e.CellStyle.BackColor = Color.LightSalmon; // light red
                        }
                        else if (value >= -2 && value <= 0)
                        {
                            e.CellStyle.BackColor = Color.LightYellow;
                        }
                        else if (value > 0)
                        {
                            e.CellStyle.BackColor = Color.LightGreen;
                        }
                    }
                }
            }

            // Make "productName" and "stickers" columns bold.
            if (dataGridView1.Columns[e.ColumnIndex].Name == "productName" ||
                dataGridView1.Columns[e.ColumnIndex].Name == "stickers")
            {
                e.CellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Bold);
            }
        }
    }
}