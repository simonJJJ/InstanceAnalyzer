﻿using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using System.Linq;
using System.Text;



namespace RoadefAnalyzer
{
    public class InstanceAnalyzer {
        public const int N = 300;
        public const string csvFilePath_batch = "_batch.csv";
        public const string csvFilePath_defects = "_defects.csv";
        public const string excelFilePath = "v2.";
        public const string OutputDir = "Analysis/Instance/";
        //public static List<batch> inputbatch = new List<batch>();
        //public static List<defects> inputdefects = new List<defects>();

        public static void Main(string[] args)
        {
            //batch[] inputbatch = new batch[N];
            //defects[] inputdefects = new defects[N];
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
                int i = 0;
                string[] s = line.Split(new char[] { ';' });
                /*inputbatch[i].Item_ID = s[0];
                inputbatch[i].Length_Item = s[1];
                inputbatch[i].Width_Item = s[2];
                inputbatch[i].Stack = s[3];
                inputbatch[i].Sequence = s[4];*/
                batch sline = new batch();
                sline.Item_ID = s[0];
                sline.Length_Item = s[1];
                sline.Width_Item = s[2];
                sline.Stack = s[3];
                sline.Sequence = s[4];
                inputbatch.Add(sline);
                i++;
            }
        }

        public static void defectsReader(string csvFilePath, List<defects> inputdefects)
        {
            string line;
            StreamReader fs = new StreamReader(csvFilePath);
            fs.ReadLine();
            while ((line = fs.ReadLine()) != null)
            {
                int i = 0;
                string[] s = line.Split(new char[] { ';' });
                /*inputdefects[i].Defect_ID = s[0];
                inputdefects[i].Plate_ID = s[1];
                inputdefects[i].X = s[2];
                inputdefects[i].Y = s[3];
                inputdefects[i].Width_Defect = s[4];
                inputdefects[i].Height_Defect = s[5];*/
                defects sline = new defects();
                sline.Defect_ID = s[0];
                sline.Plate_ID = s[1];
                sline.X = s[2];
                sline.Y = s[3];
                sline.Width_Defect = s[4];
                sline.Height_Defect = s[5];
                inputdefects.Add(sline);
                i++;
            }
        }
    }
}
