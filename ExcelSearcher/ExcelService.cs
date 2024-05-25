using System.IO;
using System;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using System.ComponentModel;

namespace ExcelSearcher
{
    public class ExcelService
    {
        public delegate void ProgressReporter(int progress, string currentFile);
        public static string[] Search(string filePath, string name, string ID, ProgressReporter progressReporter)
        {
            
            string[] files = TraverseDirectory(filePath);
            List<string> filteredFiles = new List<string>();
            if(files ==  null) { return new string[] { "无结果" }; }
            for (int i = 0;i< files.Length; i++)
            {
                progressReporter((i + 1) * 100 / files.Length, files[i]);
                if (DoContain(files[i], name, ID)) 
                {
                    filteredFiles.Add(files[i]);
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return filteredFiles.ToArray();
        }

        static string[] TraverseDirectory(string filePath)
        {
            try
            {
                return Directory.GetFiles(filePath, "*.*", SearchOption.AllDirectories);
            }
            catch(Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        static bool DoContain(string file, string name, string ID)
        {
            return SearchInExcel(file, name, ID);
        }

        static bool SearchInExcel(string file, string name, string ID)
        {
            IWorkbook workbook = null;
            FileStream fs = null;
            try
            {
                using (fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                {
                    if (Path.GetExtension(file).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = new XSSFWorkbook(fs); // for .xlsx

                    }
                    else if (Path.GetExtension(file).Equals(".xls", StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = new HSSFWorkbook(fs); // for .xls
                    }
                    else
                    {
                        //Console.WriteLine($"Unsupported file format: {file}");
                        return false;
                    }
                }

                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    ISheet sheet = workbook.GetSheetAt(i);
                    if (sheet == null)
                    {
                        continue;
                    }

                    for (int row = 0; row < sheet.LastRowNum; row++)
                    {
                        IRow currentRow = sheet.GetRow(row);
                        if (currentRow == null)
                        {
                            continue;
                        }

                        for (int cell = 0; cell < currentRow.LastCellNum; cell++)
                        {
                            ICell currentCell = currentRow.GetCell(cell);
                            if (currentCell != null && currentCell.CellType == CellType.String)
                            {
                                if (currentCell.StringCellValue.Contains(name) || currentCell.StringCellValue.Contains(ID))
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                if(fs != null)
                {
                    fs.Close();
                    fs.Dispose();
                }
                if(workbook != null)
                {
                    workbook.Close();
                    workbook.Dispose();
                }
            }
        }
    }
}
