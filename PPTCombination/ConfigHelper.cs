using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PPTCombination
{
    public class ConfigHelper
    {
        private const string SourceFileNameHead = "SourceFileName";
        private const string SourceSlideIndexHead = "SourceSlideIndex";

        public static IList<Config> ReadFromFile(string fileName)
        {
            if (fileName.EndsWith(".json"))
            {
                return ReadFromJson(fileName);
            }

            return fileName.EndsWith(".xlsx") ? ReadFromExcel(fileName) : null;
        }

        private static IList<Config> ReadFromJson(string fileName)
        {
            var jsonConfig = File.ReadAllText(fileName);
            return JsonConvert.DeserializeObject<IList<Config>>(jsonConfig.Trim());
        }

        private static IList<Config> ReadFromExcel(string fileName)
        {
            var configList = new List<Config>();

            using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                var workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheet("Sheet1");

                Config config = null;
                var index = 0;
                while (index <= sheet.LastRowNum)
                {
                    if (!FindExcelFileNameConfig(sheet, ref index, ref config))
                    {
                        break;
                    }

                    // begin to find slide info
                    if (!FindExcelSlideHead(sheet, ref index))
                    {
                        continue;
                    }

                    FindExcelSlideConfig(sheet, ref index, ref config);

                    if (config != null && config.Slides.Any())
                    {
                        configList.Add(config);
                    }
                }
            }

            return configList;
        }

        private static bool FindExcelFileNameConfig(ISheet sheet, ref int index, ref Config config)
        {
            while (index <= sheet.LastRowNum)
            {
                var row = sheet.GetRow(index++);
                if (row == null)
                {
                    continue;
                }

                var cell1 = row.GetCell(0);
                if (cell1 == null || cell1.ToString() != SourceFileNameHead)
                {
                    continue;
                }

                // SourceFileName	DestFileName
                row = sheet.GetRow(index++);
                cell1 = row.GetCell(0);
                var cell2 = row.GetCell(1);
                if (cell1 == null || cell2 == null)
                {
                    continue;
                }
                config = new Config { SourceFileName = cell1.ToString(), DestFileName = cell2.ToString(), NewFileName = (row.GetCell(2) == null ? "" : row.GetCell(2).ToString()) };

                return true;
            }

            return false;
        }

        private static bool FindExcelSlideHead(ISheet sheet, ref int index)
        {
            while (index <= sheet.LastRowNum)
            {
                var row = sheet.GetRow(index++);

                if (row == null)
                {
                    continue;
                }

                var cell1 = row.GetCell(0);

                // find next copy section
                if (cell1 != null && cell1.ToString() == SourceFileNameHead)
                {
                    index--;
                    return false;
                }

                if (cell1 == null || cell1.ToString() != SourceSlideIndexHead)
                {
                    continue;
                }

                return true;
            }

            return false;
        }

        private static void FindExcelSlideConfig(ISheet sheet, ref int index, ref Config config)
        {
            Slide slide = null;
            while (index <= sheet.LastRowNum)
            {
                var row = sheet.GetRow(index++);

                if (row == null)
                {
                    break;
                }

                var cell1 = row.GetCell(0);
                var sourceSlideIndex = -1;
                if (cell1 != null && !int.TryParse(cell1.ToString(), out sourceSlideIndex))
                {
                    index--;
                    break;
                }

                var cell2 = row.GetCell(1);
                var cell3 = row.GetCell(2);
                var cell4 = row.GetCell(3);
                var cell5 = row.GetCell(4);
                var cell6 = row.GetCell(5);

                var destSlideIndex = cell2 == null ? -1 : Convert.ToInt32(cell2.ToString());
                var sourceLeft = cell3 == null ? (float?)null : Convert.ToSingle(cell3.ToString());
                var sourceTop = cell4 == null ? (float?)null : Convert.ToSingle(cell4.ToString());
                var destLeft = cell5 == null ? (float?)null : Convert.ToSingle(cell5.ToString());
                var destTop = cell6 == null ? (float?)null : Convert.ToSingle(cell6.ToString());

                if (slide == null)
                {
                    if (sourceSlideIndex < 0 || destSlideIndex < 0)
                    {
                        continue;
                    }

                    slide = new Slide(sourceSlideIndex, destSlideIndex, sourceLeft, sourceTop, destLeft, destTop);
                }
                else
                {
                    if ((sourceSlideIndex < 0 && destSlideIndex < 0) ||
                        (slide.SourceSlideIndex == sourceSlideIndex && slide.DestSlideIndex == destSlideIndex))
                    {
                        if (sourceLeft != null || sourceTop != null || destLeft != null || destTop != null)
                            slide.Shapes.Add(new Shape(sourceLeft, sourceTop, destLeft, destTop));
                    }
                    else
                    {
                        config.Slides.Add(slide);
                        slide = new Slide(sourceSlideIndex, destSlideIndex, sourceLeft, sourceTop, destLeft, destTop);
                    }
                }
            }

            if (slide != null)
            {
                config.Slides.Add(slide);
            }
        }
    }
}
