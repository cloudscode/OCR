using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using CloudCodeOCR;
namespace CloudCodeTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string imagePath = "c:\\2.jpg";
            Console.WriteLine(System.IO.Path.GetFileNameWithoutExtension(imagePath));
            Console.WriteLine(System.IO.Path.GetFileName(imagePath));
            Console.WriteLine(System.IO.Path.GetExtension(imagePath));
            if (String.IsNullOrWhiteSpace(imagePath))
            {
               // Console.WriteLine("Usage: {0} [Path to image file]", Path.GetFileName(Assembly.GetAssembly(typeof(Program)).CodeBase));
                return;
            }

            Console.WriteLine("Running OCR for file " + imagePath);
            try
            {
                using (var ocrEngine = new OnenoteOcrEngine1())
                using (var image = Image.FromFile(imagePath))
                {
                    var imageBytes = File.ReadAllBytes(imagePath);
                    var text = ocrEngine.Recognize(imageBytes, ".jpg");
                    //var text = ocrEngine.Recognize(image, ".jpg");
                    //var text = ocrEngine.Recognize(image);
                    if (text == null)
                        Console.WriteLine("nothing recognized");
                    else {
                        lblTxt.Text = text;
                        Console.WriteLine("Recognized: " + text); }
                       
                }
            }
            catch (OcrException ex)
            {
                Console.WriteLine("OcrException:\n" + ex);
            }
            catch (Exception ex)
            {
                Console.WriteLine("General Exception:\n" + ex);
            }

            Console.ReadLine();
        }
    }
}
