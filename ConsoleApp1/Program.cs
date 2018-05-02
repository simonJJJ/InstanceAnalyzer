using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using System.Linq;
using System.Text;



namespace RoadefAnalyzer
{
    public class InstanceAnalyzer {
        public const string csvFilePath_batch = "_batch.csv";
        public const string csvFilePath_defects = "_defects.csv";
        public const string excelFilePath = "v2.";
        public const string OutputDir = "Analysis/Instance/";

        public static void Main(string[] args)
        {
            analyzeInstance();
        }

        public static void analyzeInstance()
        {
            int i = 1;
            for(i = 1; i <= 20; i++)
            {
                List<batch> inputbatch = new List<batch>();
                List<defects> inputdefects = new List<defects>();
                Directory.CreateDirectory(OutputDir);
                loadInstance(inputbatch, inputdefects, "A" + i + csvFilePath_batch, "A" + i + csvFilePath_defects);
                convertToExcelDocument(inputbatch, inputdefects, OutputDir + excelFilePath + i + ".xlsx");
            }
        }

        public static void loadInstance(List<batch> inputbatch, List<defects> inputdefects, string csvFilePath_batch, string csvFilePath_defects)
        {
            batchReader(csvFilePath_batch, inputbatch);
            defectsReader(csvFilePath_defects, inputdefects);
        }

        public static void convertToExcelDocument(List<batch> inputbatch, List<defects> inputdefects, string excelFilePath)
        {
            int i, j, row, col;
            XLWorkbook workbook = new XLWorkbook();

            #region Basis
            IXLWorksheet basisSheet = workbook.Worksheets.Add("Basis");
            row = 0;
            basisSheet.Cell(++row, 1).Value = "ItemNum";
            basisSheet.Cell(row, 2).Value = int.Parse(inputbatch[inputbatch.Count - 1].Item_ID) + 1;
            basisSheet.Cell(++row, 1).Value = "StackNum";
            basisSheet.Cell(row, 2).Value = int.Parse(inputbatch[inputbatch.Count - 1].Stack) + 1;
            basisSheet.Cell(++row, 1).Value = "DefectNum";
            basisSheet.Cell(row, 2).Value = int.Parse(inputdefects[inputdefects.Count - 1].Defect_ID) + 1;
            basisSheet.Cell(++row, 1).Value = "PlateNum";
            basisSheet.Cell(row, 2).Value = int.Parse(inputdefects[inputdefects.Count - 1].Plate_ID) + 1;
            basisSheet.SheetView.FreezeColumns(1);
            #endregion

            #region Batch
            IXLWorksheet batchSheet = workbook.Worksheets.Add("Batch");
            row = 0;
            batchSheet.Cell(++row, 1).Value = "Item_ID";
            for (i = 0, col = 2; i < inputbatch.Count(); i++, col++)
            {
                batchSheet.Cell(row, col).Value = inputbatch[i].Item_ID;
            }
            batchSheet.Cell(++row, 1).Value = "Length_Item";
            for (i = 0, col = 2; i < inputbatch.Count(); i++, col++)
            {
                batchSheet.Cell(row, col).Value = inputbatch[i].Length_Item;
            }
            batchSheet.Cell(++row, 1).Value = "Width_Item";
            for (i = 0, col = 2; i < inputbatch.Count(); i++, col++)
            {
                batchSheet.Cell(row, col).Value = inputbatch[i].Width_Item;
            }
            batchSheet.Cell(++row, 1).Value = "Stack";
            for (i = 0, col = 2; i < inputbatch.Count(); i++, col++)
            {
                batchSheet.Cell(row, col).Value = inputbatch[i].Stack;
            }
            batchSheet.Cell(++row, 1).Value = "Sequence";
            for (i = 0, col = 2; i < inputbatch.Count(); i++, col++)
            {
                batchSheet.Cell(row, col).Value = inputbatch[i].Sequence;
            }
            batchSheet.SheetView.FreezeColumns(1);
            #endregion

            #region Defects
            IXLWorksheet defectsSheet = workbook.Worksheets.Add("Defects");
            row = 0;
            defectsSheet.Cell(++row, 1).Value = "Defect_ID";
            for (j = 0, col = 2; j < inputdefects.Count(); j++, col++)
            {
                defectsSheet.Cell(row, col).Value = inputdefects[j].Defect_ID;
            }
            defectsSheet.Cell(++row, 1).Value = "Plate_ID";
            for (j = 0, col = 2; j < inputdefects.Count(); j++, col++)
            {
                defectsSheet.Cell(row, col).Value = inputdefects[j].Plate_ID;
            }
            defectsSheet.Cell(++row, 1).Value = "X";
            for (j = 0, col = 2; j < inputdefects.Count(); j++, col++)
            {
                defectsSheet.Cell(row, col).Value = inputdefects[j].X;
            }
            defectsSheet.Cell(++row, 1).Value = "Y";
            for (j = 0, col = 2; j < inputdefects.Count(); j++, col++)
            {
                defectsSheet.Cell(row, col).Value = inputdefects[j].Y;
            }
            defectsSheet.Cell(++row, 1).Value = "Width_Defect";
            for (j = 0, col = 2; j < inputdefects.Count(); j++, col++)
            {
                defectsSheet.Cell(row, col).Value = inputdefects[j].Width_Defect;
            }
            defectsSheet.Cell(++row, 1).Value = "Height_Defect";
            for (j = 0, col = 2; j < inputdefects.Count(); j++, col++)
            {
                defectsSheet.Cell(row, col).Value = inputdefects[j].Height_Defect;
            }
            defectsSheet.SheetView.FreezeColumns(1);
            #endregion

            for (i = 1; i <= workbook.Worksheets.Count; ++i)
            {
                workbook.Worksheets.Worksheet(i).Column(1).AdjustToContents();
            }
            workbook.SaveAs(excelFilePath);

        }

        public struct batch
        {
            public string Item_ID;
            public string Length_Item;
            public string Width_Item;
            public string Stack;
            public string Sequence;
        }

        public struct defects
        {
            public string Defect_ID;
            public string Plate_ID;
            public string X;
            public string Y;
            public string Width_Defect;
            public string Height_Defect;
        }

        public static void batchReader(string csvFilePath, List<batch> inputbatch)
        {
            string line;
            StreamReader fs = new StreamReader(csvFilePath);
            fs.ReadLine();
            while ((line = fs.ReadLine()) != null)
            {
                string[] s = line.Split(new char[] { ';' });
                batch sline = new batch();
                sline.Item_ID = s[0];
                sline.Length_Item = s[1];
                sline.Width_Item = s[2];
                sline.Stack = s[3];
                sline.Sequence = s[4];
                inputbatch.Add(sline);
            }
        }

        public static void defectsReader(string csvFilePath, List<defects> inputdefects)
        {
            string line;
            StreamReader fs = new StreamReader(csvFilePath);
            fs.ReadLine();
            while ((line = fs.ReadLine()) != null)
            {
                string[] s = line.Split(new char[] { ';' });
                defects sline = new defects();
                sline.Defect_ID = s[0];
                sline.Plate_ID = s[1];
                sline.X = s[2];
                sline.Y = s[3];
                sline.Width_Defect = s[4];
                sline.Height_Defect = s[5];
                inputdefects.Add(sline);
            }
        }
    }
}
