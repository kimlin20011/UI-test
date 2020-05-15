using System;
using System.IO;

//這個檔案主要是讀取輸入與接收的參數，並做進一步的
namespace SkypeAutoCapture
{
    class Program
    {
        static void Main(string[] args)
        {
            //檢查屬性參數
            if (args.Length >= 2)
            {
                SkypeSession skypeSession = new SkypeSession();
                skypeSession.Setup();

                //設定檔案路徑
                String folderPath = args[1] + "\\Skype";  //建立SKYPE資料夾
                int folderName = 0;

                //system 的檔案模組
                Directory.CreateDirectory(folderPath);

                while (Directory.Exists(folderPath + "\\" + Convert.ToString(folderName)))
                {
                    folderName++;
                }

                //進去划功能
                if ("0" == args[0])
                {
                    Directory.CreateDirectory(folderPath + "\\" + Convert.ToString(folderName));
                    if (args.Length >= 4)
                    {
                        skypeSession.autoScrollScreenShotOnePage(folderPath + "\\" + Convert.ToString(folderName), int.Parse(args[2]), int.Parse(args[3]));
                    }
                    else if (args.Length >= 3)
                    {
                        skypeSession.autoScrollScreenShotOnePage(folderPath + "\\" + Convert.ToString(folderName), int.Parse(args[2]), Int32.MaxValue);
                    }
                }
                else if ("1" == args[0])
                {
                    skypeSession.autoClickContactPersonScrollScreenShotEachPage(folderPath, folderName);
                }

                skypeSession.TearDown();
            }
        }
    }
}
