using System;
using Corel.Interop.VGCore;

namespace TESTCoreldraw
{
    class Program
    {
        public static void Main(String[] args)
        {
            //源文件
            String source = @"C:\Users\Yao\Documents\Corel\world.cdr";

            //输出文件
            String release = @"C:\Users\Yao\Documents\Corel\world2.dxf";

            //String release = @"C:\Users\Yao\Documents\Corel\world2.ai";

            //String release = @"C:\Users\Yao\Documents\Corel\world2.pdf";

            //String release = @"C:\Users\Yao\Documents\Corel\world2.jpg";

            //String release = @"C:\Users\Yao\Documents\Corel\world2.png";


            //cdr转换成pdf
            //new FileConverter().cdr2PDF(source, release);


            //cdr转换成png
            new FileConverter().cdr2ExportBitmap(source, release);



            //创建并打开新的.cdr文档
            //try
            //{
            //    CreateDocument.CreateTextInCorelDRAW("Hello, B**ches", "Arial", 36.0f);
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("Error occurred: {0}", ex.Message);
            //}



        }



      


    }
}