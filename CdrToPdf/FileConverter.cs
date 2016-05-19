using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Corel.Interop.VGCore;

/**********************************************************************************     
     * Created by  Yao  on  5/20/2016 12:04:33 AM     
     * README:.cdr文件类型转换
     * ============================================================================
     * CHANGELOG：
    ********************************************************************************/
namespace TESTCoreldraw
{
    public class FileConverter
    {
        /// <summary>
        /// 输出pdf
        /// </summary>
        /// <param name="sourceFilePath"></param>
        /// <param name="outputFilePath"></param>
        public void cdr2PDF(string sourceFilePath, string outputFilePath)
        {

            Application application = new Application();

            Document document = (Document)application.OpenDocument(sourceFilePath, 1);

            document.PublishToPDF(outputFilePath);

            document.Close();
        }

        /// <summary>
        /// 导出多种格式
        /// </summary>
        /// <param name="sourceFilePath"></param>
        /// <param name="outputFilePath"></param>
        public void cdr2ExportBitmap(string sourceFilePath, string outputFilePath)
        {

            Application application = new Application();

            Document document = application.OpenDocument(sourceFilePath, 1);


            //.cdr转换成.png,.jpg,.ai,.dxf,.dwg对应的cdrFilter依次为
            // cdrPNG, cdrJPEG, cdrAI, cdrDXF, cdrDWG
            //其他参考cdrTIFF, cdrPSD, cdrDOC, cdrSVG, cdrEPS 
            document.ExportBitmap(
                     outputFilePath,
                     Corel.Interop.VGCore.cdrFilter.cdrDXF,
                     Corel.Interop.VGCore.cdrExportRange.cdrCurrentPage,
                     Corel.Interop.VGCore.cdrImageType.cdrRGBColorImage,
                     0, 0, 72, 72,
                     Corel.Interop.VGCore.cdrAntiAliasingType.cdrNoAntiAliasing,
                     false,
                     true,
                     true,
                     false,
                     Corel.Interop.VGCore.cdrCompressionType.cdrCompressionNone,
                     null).Finish();

            //输出PNG
            //document.ExportBitmap(
            //        outputFilePath,
            //        Corel.Interop.VGCore.cdrFilter.cdrPNG,
            //        Corel.Interop.VGCore.cdrExportRange.cdrCurrentPage,
            //        Corel.Interop.VGCore.cdrImageType.cdrRGBColorImage,
            //        0, 0, 72, 72,
            //        Corel.Interop.VGCore.cdrAntiAliasingType.cdrNoAntiAliasing,
            //        false,
            //        true,
            //        true,
            //        false,
            //        Corel.Interop.VGCore.cdrCompressionType.cdrCompressionNone,
            //        null).Finish();


            document.Close();

        }





    }
}