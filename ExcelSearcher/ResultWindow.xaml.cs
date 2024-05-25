using System;
using System.Collections.Generic;
using System.Windows;

namespace ExcelSearcher
{
    /// <summary>
    /// Interaction logic for ResultWindow.xaml
    /// </summary>
    public partial class ResultWindow : Window
    {
        private List<ResultItem> resultList;
        public ResultWindow(IEnumerable<string> results)
        {
            InitializeComponent();
            this.Closed += ResultWindow_Closed;
            resultList = new List<ResultItem>();
            foreach (string result in results)
            {
                resultList.Add(new ResultItem { Result = result });
            }

            // 将集合作为 DataGrid 的数据源
            ResultDataGrid.ItemsSource = resultList;
        }

        private void ResultWindow_Closed(object sender, EventArgs e)
        {
            resultList.Clear();
            resultList = null;

            this.Closed -= ResultWindow_Closed;
        }
    }

    public class ResultItem
    {
        public string Result { get; set; }
    }

    
}
