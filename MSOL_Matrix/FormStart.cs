using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MSOL_Matrix
{
    public partial class FormStart : Form
    {
        public FormStart()
        {
            InitializeComponent();
        }

        Stopwatch stopwatch;
        private async void button1_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            this.StartBtn.Enabled = false;
             stopwatch = new Stopwatch();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileInfo excelFileInfo = new FileInfo(openFileDialog.FileName);
                string fileReadResult = string.Empty;
                DataTable ExcelDt = ExcelReader.tryReadExcel(excelFileInfo, fileReadResult);

                if (string.IsNullOrWhiteSpace(fileReadResult))
                {
                    labelWatch.Visible = true;
                   
                    stopwatch.Start();
                    timer1.Enabled = true;

                    DataTablesModel dataTablesModel = await Task.Run(() => MatrixCalculator.calculateRawMatrix(ExcelDt));
                    await Task.Run(() => ExcelWriter.saveRawDtAndMatrixToExcel(excelFileInfo.DirectoryName, ExcelDt, dataTablesModel));
                    stopwatch.Stop();
                    MessageBox.Show("Operations Complete", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    stopwatch.Stop();
                    MessageBox.Show($"Error occured: {fileReadResult}", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            this.StartBtn.Enabled = true;
            openFileDialog.Dispose();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            labelWatch.Text = string.Format("{0:hh\\:mm\\:ss\\.}", stopwatch.Elapsed);
        }
    }
}
