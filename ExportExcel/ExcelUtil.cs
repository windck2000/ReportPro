using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;

namespace ExportExcel
{
    // Token: 0x02000005 RID: 5
    internal class ExcelUtil
    {
        // Token: 0x06000004 RID: 4 RVA: 0x00002443 File Offset: 0x00000643
        public ExcelUtil()
        {
            this.hssfWorkbook = new XSSFWorkbook();
        }

        // Token: 0x06000005 RID: 5 RVA: 0x00002456 File Offset: 0x00000656
        public ExcelUtil(string company, string subject)
        {
        }

        // Token: 0x06000006 RID: 6 RVA: 0x00002460 File Offset: 0x00000660
        public void CPKData(DataTable dtResult, int dataStartLoc, int snCount)
        {
            TitleLoc titleLoc = new TitleLoc(snCount);
            string[] array = new string[]
            {
                "MAX",
                "MIN",
                "AVG",
                "STD",
                "Cpu",
                "Cpl",
                "Cp > 1",
                "Ca < 1",
                "Cpk > 1",
                "Result"
            };
            XSSFSheet xssfsheet = (XSSFSheet)this.hssfWorkbook.CreateSheet("Resutl");
            XSSFCellStyle xssfcellStyle = (XSSFCellStyle)this.hssfWorkbook.CreateCellStyle();
            xssfcellStyle.FillBackgroundColor = 10;
            xssfcellStyle.FillPattern = FillPattern.SolidForeground;
            for (int i = 0; i < dtResult.Rows.Count; i++)
            {
                XSSFRow xssfrow = (XSSFRow)xssfsheet.CreateRow(i);
                TitleTYpe titleTYpe = (TitleTYpe)Array.IndexOf<string>(array, dtResult.Rows[i][0].ToString());
                for (int j = 0; j < dtResult.Columns.Count; j++)
                {
                    xssfrow.CreateCell(j);
                    if (j > 0 && i >= dataStartLoc - 1 && i < dataStartLoc + snCount + 2)
                    {
                        xssfrow.GetCell(j).SetCellType(CellType.Numeric);
                        double cellValue = 0.0;
                        double.TryParse(dtResult.Rows[i][j].ToString(), out cellValue);
                        xssfrow.GetCell(j).SetCellValue(cellValue);
                    }
                    else if (titleTYpe == (TitleTYpe)(-1) || Array.IndexOf<string>(array, dtResult.Rows[i][j].ToString()) > -1)
                    {
                        xssfrow.GetCell(j).SetCellValue(dtResult.Rows[i][j].ToString());
                    }
                    else
                    {
                        string cellFormula = titleLoc.getCellFormula(titleTYpe, j);
                        if (!string.IsNullOrEmpty(cellFormula))
                        {
                            xssfrow.GetCell(j).SetCellType(CellType.Numeric);
                            xssfrow.GetCell(j).SetCellFormula(cellFormula);
                        }
                    }
                }
            }
            string empty = string.Empty;
        }

        // Token: 0x06000007 RID: 7 RVA: 0x0000267C File Offset: 0x0000087C
        public void WriteToFile(ExportType exportType)
        {

            string path = "abc.xls";
            if (exportType == ExportType.FileDialog)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.DefaultExt = "xlsx";
                saveFileDialog.Filter = "All Files(*.xlsx)|*.xlsx";
                if (!(bool)saveFileDialog.ShowDialog())
                {
                    return;
                }
                path = saveFileDialog.FileName;
            }
            FileStream fileStream = new FileStream(path, FileMode.Create);
            this.hssfWorkbook.Write(fileStream);
            fileStream.Close();
        }

        // Token: 0x0400001E RID: 30
        private XSSFWorkbook hssfWorkbook;
    }
}
