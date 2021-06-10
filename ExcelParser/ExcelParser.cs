using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Spire.Xls;

namespace ExcelParser
{
    public class ExcelParser : IExcelParser
    {
        public Workbook WorkBook = null;

        public void CopyRangeToNewWorkbook(Worksheet oldWs, Worksheet newWs, string leftUpCornerOnOldFile, string rightDownCornerOnOldFile, string leftUpCornerOnNewFile, string rightDownOnOldFile)
        {
            CellRange range = ;

        }

        public int GetColumnCount(Worksheet ws)
        {
            return ws.LastColumn + 1;
        }

        public List<string> GetContentOfColumn(Worksheet ws, int columnIndex, bool includeHeader)
        {
            int rowStartIndex = 0;
            if (includeHeader)
            {
                rowStartIndex = 1;
            }

            List<string> ret = new List<string>();
            for (int i = rowStartIndex; i < GetRowCount(ws); i++)
            {
                string value = ws.GetText(i, columnIndex);
                ret.Add(value);
            }

            return ret;
        }

        public List<string> GetContentOfRow(Worksheet ws, int rowIndex)
        {
            //int columnStartIndex = 1;            

            List<string> ret = new List<string>();
            for (int i = 1; i < GetColumnCount(ws); i++)
            {
                string value = ws.GetText(rowIndex, i);
                ret.Add(value);
            }

            return ret;
        }

        public List<string> GetHeads(Worksheet ws)
        {
            return GetContentOfRow(ws, ws.FirstRow);
        }

        public string[,] GetRangeData(Worksheet ws, int leftCornerRowIndex, int leftCornerColumnIndex, int rightCornerRowIndex, int rightCornerColumnIndex)
        {
            string[,] ret = new string[rightCornerRowIndex - leftCornerRowIndex, rightCornerColumnIndex - leftCornerColumnIndex];            

            for (int i = leftCornerRowIndex; i < rightCornerColumnIndex; i++)
            {
                for (int j = leftCornerRowIndex; j < rightCornerRowIndex; j++)
                {
                   ret[j,i]  = ws.GetText(j, i);
                }
            }

            return ret;
        }

        public int GetRowCount(Worksheet ws)
        {
            return ws.LastRow + 1;
        }

        public Worksheet GetWorksheet(Workbook wb, int index)
        {
            return wb.Worksheets[index];
        }

        public int GetWorksheetCount(Worksheet ws)
        {
            return WorkBook.Worksheets.Count();
        }

        public Workbook LoadExcel(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new Exception("文件路径错误！");                
            }

            Workbook wb = new Workbook();
            try
            {
                wb.LoadFromFile(filePath, ExcelVersion.Version97to2003);
            }
            catch (Exception ex)
            {
                MessageBox.Show("加载文件失败，原因：" + ex); ;
            }
            

            return wb;
        }
    }
}
