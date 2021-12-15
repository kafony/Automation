using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.PowerPointApi.Tools;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace PPTCombination
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var configList = ReadFromJson();

                DoCombination(configList);

                Exit();
            }
            catch (Exception e)
            {
                Console.WriteLine("程序异常，请检查配置文件！\n\n");
                Console.WriteLine(e);
                Console.WriteLine("\n\n按任意键退出");
                Console.ReadKey();
            }
        }

        static IList<Config> ReadFromJson()
        {
            var jsonConfig = File.ReadAllText("config.json");
            return JsonConvert.DeserializeObject<IList<Config>>(jsonConfig.Trim());
        }

        static void DoCombination(IList<Config> configList)
        {
            if (configList == null || configList.Count == 0 || configList.All(d => d.Slides == null || d.Slides.Count == 0))
            {
                Console.WriteLine("没有需要处理的文件\n");
                return;
            }

            foreach (var config in configList)
            {
                if (config.Slides == null || config.Slides.Count == 0)
                {
                    continue;
                }

                Console.WriteLine($"正在处理 {config.SourceFileName}.pptx => {config.DestFileName}.pptx ...\n");

                var sourcePPT = new PPTApplication();
                var sourcePresentation = sourcePPT.GetPresentation(config.SourceFileName);

                var destPPT = new PPTApplication();
                var destPresentation = destPPT.GetPresentation(config.DestFileName);

                var step = 0;
                foreach (var slide in config.Slides.OrderBy(s => s.DestSlideIndex))
                {
                    var sourceSlide = sourcePresentation.Slides.Range(slide.SourceSlideIndex);

                    // copy slide
                    sourceSlide.Copy();
                    destPresentation.Slides.Paste(slide.DestSlideIndex + step);
                    step++;
                }

                destPPT.Save(config.DestFileName + "new");

                sourcePPT.Dispose();
                destPPT.Dispose();
            }
        }

        static void Exit()
        {
            Console.WriteLine("处理完成，5s后自动关闭。");
            Thread.Sleep(5 * 1000);
            Environment.Exit(0);
        }
    }
}