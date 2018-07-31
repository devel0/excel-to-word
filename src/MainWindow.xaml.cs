using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Xceed.Words.NET;

namespace ExcelToWord
{

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += MainWindow_Loaded;
            this.Closing += MainWindow_Closing;
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {            
            config.ExcelPathfilename = ExcelPathfilename.Text;
            config.WordPathfilename = WordPathfilename.Text;
            config.OutputFolder = OutputFolder.Text;

            config.Save();
        }

        internal Global global { get { return Global.Instance; } }
        public Config config { get { return global.Config; } }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            dg.ItemsSource = config.Mapping;
            ExcelPathfilename.Text = config.ExcelPathfilename;
            WordPathfilename.Text = config.WordPathfilename;
            OutputFolder.Text = config.OutputFolder;
        }

        private void AddMapping(object sender, RoutedEventArgs e)
        {
            config.Mapping.Add(new ExcelToWordConfigItem());
        }

        private void RemoveMapping(object sender, RoutedEventArgs e)
        {
            if (dg.SelectedItems.Count > 0)
            {
                config.Mapping.Remove(dg.SelectedItem as ExcelToWordConfigItem);
            }
        }

        private void Generate(object sender, RoutedEventArgs e)
        {
#if !DEBUG
            try
#endif
            {
                if (ExcelPathfilename.Text.Trim().Length == 0) { MessageBox.Show($"Please specify Excel datasource Pathfilename"); return; }
                if (WordPathfilename.Text.Trim().Length == 0) { MessageBox.Show($"Please specify Word template Pathfilename"); return; }
                if (OutputFolder.Text.Trim().Length == 0) { MessageBox.Show($"Please specify output directory"); return; }

                if (!File.Exists(ExcelPathfilename.Text)) { MessageBox.Show($"Unable to find [{ExcelPathfilename}]"); return; }
                if (!File.Exists(WordPathfilename.Text)) { MessageBox.Show($"Unable to find [{WordPathfilename}]"); return; }
                if (!Directory.Exists(OutputFolder.Text)) { MessageBox.Show($"Unable to find [{OutputFolder}] folder"); return; }

                var colNameToColIdx = new Dictionary<string, int>();

                using (var wb = new ClosedXML.Excel.XLWorkbook(ExcelPathfilename.Text))
                {
                    using (var ws = wb.Worksheets.First())
                    {
                        var colCnt = ws.ColumnsUsed().Count();
                        var rowCnt = ws.RowsUsed().Count();

                        for (int col = 1; col <= colCnt; ++col)
                        {
                            var cellValue = ws.Cell(1, col).Value;
                            if (!(cellValue is string)) continue;
                            var cellText = (string)cellValue;

                            var q = config.Mapping.FirstOrDefault(w => w.ColumnName.ToUpper() == cellText.ToUpper());
                            if (q != null) colNameToColIdx.Add(q.ColumnName, col);
                        }

                        for (int row = 2; row <= rowCnt; ++row)
                        {
                            var suffix = string.Format("{0:0000}", row);
                            var outputFilename = Path.Combine(
                                OutputFolder.Text,
                                Path.GetFileNameWithoutExtension(WordPathfilename.Text) + $"-{suffix}.docx");

                            File.Copy(WordPathfilename.Text, outputFilename, true);

                            var docx = DocX.Load(outputFilename);

                            foreach (var x in colNameToColIdx)
                            {
                                var columnName = x.Key;
                                var colIndex = x.Value;

                                var token = config.Mapping.First(w => w.ColumnName == columnName).TokenToReplace;

                                var valueToInsert = ws.Cell(row, colIndex).Value;
                                if (valueToInsert as string == null) continue;
                                docx.ReplaceText(token, (string)valueToInsert);
                            }

                            docx.Save();
                        }
                    }
                }
            }
#if !DEBUG
            catch (Exception ex)
            {
                MessageBox.Show($"unhandled exception : {ex.Message}");
            }
#endif

        }

    }

}
