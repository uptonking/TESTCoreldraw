using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Corel.Interop.VGCore;

namespace TESTCoreldraw
{
    class CreateDocument
    {
       public  static void CreateTextInCorelDRAW(string text, string fontName,
                                       float fontSize)
        {
            Type pia_type = Type.GetTypeFromProgID("CorelDRAW.Application.17");
            Application app = Activator.CreateInstance(pia_type) as Application;
            app.Visible = true;
            Document doc = app.ActiveDocument;
            if (doc == null)
                doc = app.CreateDocument();
            Shape shape = doc.ActiveLayer.CreateArtisticText(
              0.0, 0.0, text, cdrTextLanguage.cdrLanguageMixed,
              cdrTextCharSet.cdrCharSetMixed, fontName, fontSize,
              cdrTriState.cdrUndefined, cdrTriState.cdrUndefined,
              cdrFontLine.cdrMixedFontLine, cdrAlignment.cdrLeftAlignment);
        }

        //static void Main(string[] args)
        //{
        //    try
        //    {
        //        CreateTextInCorelDRAW("Hello, B**ches", "Arial", 36.0f);
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Error occurred: {0}", ex.Message);
        //    }
        //}
    }
}
