using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;

namespace ExcelParser
{
    public interface IExcelParser
    {
        /// <summary>
        /// 获取Excel工作簿
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        Workbook LoadExcel(string filePath);
        /// <summary>
        /// 获取Worksheet数量
        /// </summary>
        /// <param name="wb"></param>
        /// <returns></returns>
        int GetWorksheetCount(Worksheet ws);
        /// <summary>
        /// 获取指定的Worksheet
        /// </summary>
        /// <returns></returns>
        Worksheet GetWorksheet(Workbook wb, int index);
        /// <summary>
        /// 获取列数量
        /// </summary>
        /// <param name="w"></param>
        /// <returns></returns>
        int GetColumnCount(Worksheet ws);
        /// <summary>
        /// 获取行数量
        /// </summary>
        /// <param name="wb"></param>
        /// <returns></returns>
        int GetRowCount(Worksheet ws);
        /// <summary>
        /// 获取所有表头
        /// </summary>
        /// <param name="wb"></param>
        /// <returns></returns>
        List<string> GetHeads(Worksheet ws);
        /// <summary>
        /// 获取整列数据
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="columnIndex"></param>
        /// <param name="includeHeader">是否包含表头信息</param>
        /// <returns></returns>
        List<string> GetContentOfColumn(Worksheet ws, int columnIndex, bool includeHeader);
        /// <summary>
        /// 获取整行数据
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        List<string> GetContentOfRow(Worksheet ws, int rowIndex);
        /// <summary>
        /// 获取指定范围的表数据
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="leftCornerX"></param>
        /// <param name="leftCornerY"></param>
        /// <param name="rightCornerX"></param>
        /// <param name="rightCornerY"></param>
        /// <returns></returns>
        string[,] GetRangeData(Worksheet ws, int leftCornerRowIndex, int leftCornerColumnIndex, int rightCornerRowIndex, int rightCornerColumnIndex);

        //void CreateNewWorkbook(string newFilePath);
        //void CopyRowToNewWorkbook(int RowIndexOnOldFile, int RowIndexOnNewFile);
        //void CopyColumnToNewWorkbook(int ColumnIndexOnOldFile, int ColumnIndexOnNewFile);
        void CopyRangeToNewWorkbook(string leftUpCornerOnOldFile, string rightDownCornerOnOldFile, string leftUpCornerOnNewFile, string rightDownOnOldFile);        
    }
}
