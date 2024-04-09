using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace BarcodeGenerator
{
    public static class ExcelHelper
    {
        private static XSSFCell GetCell(ISheet sheet,int row, int col)
        {
            var tempRow = sheet.GetRow(row) as XSSFRow ?? sheet.CreateRow(row) as XSSFRow;
            var cell = tempRow.GetCell(col) as XSSFCell ?? tempRow.CreateCell(col) as XSSFCell;
            return cell;
        }

        /// <summary>
        /// 导出条码
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="barcodeList"></param>
        /// <param name="title"></param>
        public static void ExportBarcodeList(string fileName, List<string> barcodeList, string title)
        {
            //FileStream fs = File.Create(fileName);
            //IWorkbook xssfworkbook = new XSSFWorkbook();
            // var sheet = xssfworkbook.CreateSheet("sheet0");

            //GetCell(sheet,0, 0).SetCellValue(title);

            //for (int i = 0; i < barcodeList.Count; i++)
            //{
            //    GetCell(sheet, i + 1, 0).SetCellValue(barcodeList[i]);
            //}

            //sheet.ForceFormulaRecalculation = true;
            //xssfworkbook.Write(fs);
            //xssfworkbook.Close();
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                IWorkbook xssfworkbook = new XSSFWorkbook();
                var sheet = xssfworkbook.CreateSheet("sheet0");

                GetCell(sheet, 0, 0).SetCellValue(title);

                for (int i = 0; i < barcodeList.Count; i++)
                {
                    GetCell(sheet, i + 1, 0).SetCellValue(barcodeList[i]);
                }

                xssfworkbook.Write(fs);
                xssfworkbook.Close();
            }
        }
    }
}
