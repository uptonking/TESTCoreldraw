using System;
using Corel.Interop.CorelDRAW;

namespace CdrToPdf
{
	class Program
	{
		public static void Main(String[] args)
		{
            String source = @"C:\Users\Yao\Documents\Corel\iphone.cdr";
			
			String release = @"C:\Users\Yao\Documents\Corel\new.pdf";
			
			Application application = new Application();
			
			Document document = (Document)application.OpenDocument(source, 1);
			
			document.PublishToPDF(release);
			
			document.Close();
		}
	}
}