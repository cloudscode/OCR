using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using System.Xml.XPath;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using System.Collections.Generic;

namespace CloudCodeOCR
{
    public sealed class OnenoteOcrEngine1
        : IOcrEngine, IDisposable
    {
        private readonly OnenotePages _page;
        public static Application _app = null;
        private const int PollAttempts = 200;
        private const int PollInterval = 250;

        public OnenoteOcrEngine1()
        {
            if (_app == null)
            {
               // _app = new ApplicationClass();
                _app = this.Try(() => new ApplicationClass(), e => new OcrException("Error initializing OneNote", e));
            }          
            this._page = this.Try(() => new OnenotePages(_app), e => new OcrException("Error initializing pages.", e));
        }

        public void Dispose()
        {
            if (this._page != null)
            {
                this._page.Delete();
            }
        }

        public string Recognize(Image image)
        {
            return this.RecognizeIntern(image);
        }
        public string Recognize(string imagePath)
        {
            return Recognize(Image.FromFile(imagePath), System.IO.Path.GetExtension(imagePath));
        }
        public string Recognize(byte[] imagedata, string extention)
        {
            this._page.CreateImageTagExtend(imagedata, extention);
            int num = 0;
            do
            {
                Thread.Sleep(PollInterval);
                this._page.Reload();
                string str = this._page.ReadOcrText();
                if (str != null)
                {
                    return str;
                }
            }
            while (num++ < 200);
            return "";
        }

        public string Recognize(Image image, string extention)
        {
            return this._page.fnOCR(image, extention);
        }

        private string RecognizeIntern(Image image)
        {
            return null;
        }

        private T Try<T>(Func<T> action, Func<Exception, Exception> excecption)
        {
            T local;
            try
            {
                local = action();
            }
            catch (Exception exception)
            {
                throw excecption(exception);
            }
            return local;
        }

    }

    internal sealed class OnenotePages
    {
        private readonly Application _app;
        private XDocument _document;
        private string _pageId;
        private const string DefaultOutline = "<one:Outline xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\"><one:OEChildren><one:OE><one:T><![CDATA[A]]></one:T></one:OE></one:OEChildren></one:Outline>";
        private string ns;
        private XDocument onenotedoc;
        private const string OneNoteNamespace = "http://schemas.microsoft.com/office/onenote/2013/onenote";

        public OnenotePages(Application _app)
        {
            string str;
            string str2;
            this._app = _app;
            _app.OpenHierarchy(AppDomain.CurrentDomain.BaseDirectory + "temp.one", string.Empty, out str, CreateFileType.cftSection);
            this._app.CreateNewPage(str, out this._pageId);
            _app.GetHierarchy(string.Empty, HierarchyScope.hsPages, out str2);
            this.onenotedoc = XDocument.Parse(str2);
        }

        public void AddImage(Image image)
        {
            XElement content = this.CreateImageTag(image);
            this.Oe.Add(content);
        }

        public void Clear()
        {
            this.Oe.Elements().Remove<XElement>();
        }

        public int CompareEndCharPositionByLine(EndCharPosition src, EndCharPosition tgt)
        {
            if (src.line < tgt.line)
            {
                return -1;
            }
            if (src.line == tgt.line)
            {
                return 0;
            }
            return 1;
        }

        private XElement CreateImageTag(Image image)
        {
            XElement element = new XElement(XName.Get("Image", "http://schemas.microsoft.com/office/onenote/2013/onenote"));
            XElement content = new XElement(XName.Get("Data", "http://schemas.microsoft.com/office/onenote/2013/onenote"))
            {
                Value = this.ToBase64(image)
            };
            element.Add(content);
            return element;
        }

        public void CreateImageTagExtend(byte[] imagedata, string extension)
        {
            
            int num;
            string content = Convert.ToBase64String(imagedata);
            XNamespace namespace2 = this.onenotedoc.Root.Name.Namespace;
            string str2 = this._pageId;
            string extensionType = extension.ToLower().Substring(1);
           
            if (extensionType != null )
            {
                if ((extensionType == "jpg") || (extensionType == "png") || (extensionType == "emf") || (extensionType == "auto"))
                {
                    num = 100;
                    int num2 = 100;
                    XDocument document2 = new XDocument(new object[] { new XElement((XName)(namespace2 + "Page"), 
                        new XElement((XName)(namespace2 + "Outline"), 
                            new XElement((XName)(namespace2 + "OEChildren"), 
                                new XElement((XName)(namespace2 + "OE"), 
                                    new XElement((XName)(namespace2 + "Image"), 
                                        new object[] { new XAttribute("format", extensionType),
                                            new XAttribute("originalPageNumber", "0"), 
                                            new XElement((XName)(namespace2 + "Position"),
                                                new object[] { new XAttribute("x", "0"), 
                                                    new XAttribute("y", "0"), 
                                                    new XAttribute("z", "0") }), 
                                                    new XElement((XName)(namespace2 + "Size"),
                                                        new object[] { new XAttribute("width", num.ToString()), 
                                                            new XAttribute("height", num2.ToString()) }), 
                                                            new XElement((XName)(namespace2 + "Data"), 
                                                                content) }))))) });
                    document2.Root.SetAttributeValue("ID", this._pageId);
                    this._app.UpdatePageContent(document2.ToString(), DateTime.MinValue);
                    this._app.NavigateTo(this._pageId);
                   
                }
            }  
        }

        public void Delete()
        {
            this._app.DeleteHierarchy(this._pageId);
        }

        private void fnOCR(string v_strImgPath)
        {
            FileInfo info = new FileInfo(v_strImgPath);
            using (MemoryStream stream = new MemoryStream())
            {
                Bitmap bitmap = new Bitmap(v_strImgPath);
                switch (info.Extension.ToLower())
                {
                    case ".jpg":
                        bitmap.Save(stream, ImageFormat.Jpeg);
                        break;

                    case ".jpeg":
                        bitmap.Save(stream, ImageFormat.Jpeg);
                        break;

                    case ".gif":
                        bitmap.Save(stream, ImageFormat.Gif);
                        break;

                    case ".bmp":
                        bitmap.Save(stream, ImageFormat.Bmp);
                        break;

                    case ".tiff":
                        bitmap.Save(stream, ImageFormat.Tiff);
                        break;

                    case ".png":
                        bitmap.Save(stream, ImageFormat.Png);
                        break;

                    case ".emf":
                        bitmap.Save(stream, ImageFormat.Emf);
                        break;

                    default:
                        return;
                }
                string content = Convert.ToBase64String(stream.GetBuffer());
                XNamespace namespace2 = this.onenotedoc.Root.Name.Namespace;
                XElement element = this.onenotedoc.Descendants((XName)(namespace2 + "Page")).FirstOrDefault<XElement>();
                string str2 = element.Attribute("ID").Value;
                if (element == null)
                {
                    return;
                }
                string extensionType = info.Extension.ToLower().Substring(1);
                if (extensionType == null)
                {
                    goto Label_0209;
                }
                if (!(extensionType == "jpg"))
                {
                    if (extensionType == "png")
                    {
                        goto Label_01F7;
                    }
                    if (extensionType == "emf")
                    {
                        goto Label_0200;
                    }
                    goto Label_0209;
                }
                string str3 = "jpg";
                goto Label_0212;
            Label_01F7:
                str3 = "png";
                goto Label_0212;
            Label_0200:
                str3 = "emf";
                goto Label_0212;
            Label_0209:
                str3 = "auto";
            Label_0212: ;
                XDocument document = new XDocument(new object[] { new XElement((XName)(namespace2 + "Page"), new XElement((XName)(namespace2 + "Outline"), new XElement((XName)(namespace2 + "OEChildren"), new XElement((XName)(namespace2 + "OE"), new XElement((XName)(namespace2 + "Image"), new object[] { new XAttribute("format", str3), new XAttribute("originalPageNumber", "0"), new XElement((XName)(namespace2 + "Position"), new object[] { new XAttribute("x", "0"), new XAttribute("y", "0"), new XAttribute("z", "0") }), new XElement((XName)(namespace2 + "Size"), new object[] { new XAttribute("width", bitmap.Width.ToString()), new XAttribute("height", bitmap.Height.ToString()) }), new XElement((XName)(namespace2 + "Data"), content) }))))) });
                document.Root.SetAttributeValue("ID", str2);
                this._app.UpdatePageContent(document.ToString(), DateTime.MinValue);
                this._app.NavigateTo(str2);
            }
        }

        public string fnOCR(Image img, string extension)
        {
            string str7;
            using (MemoryStream stream = new MemoryStream())
            {
                string str2;
                Func<XNode, bool> predicate = null;
                Bitmap bitmap = new Bitmap(img);
                switch (extension.ToLower())
                {
                    case ".jpg":
                        bitmap.Save(stream, ImageFormat.Jpeg);
                        break;

                    case ".jpeg":
                        bitmap.Save(stream, ImageFormat.Jpeg);
                        break;

                    case ".gif":
                        bitmap.Save(stream, ImageFormat.Gif);
                        break;

                    case ".bmp":
                        bitmap.Save(stream, ImageFormat.Bmp);
                        break;

                    case ".tiff":
                        bitmap.Save(stream, ImageFormat.Tiff);
                        break;

                    case ".png":
                        bitmap.Save(stream, ImageFormat.Png);
                        break;

                    case ".emf":
                        bitmap.Save(stream, ImageFormat.Emf);
                        break;

                    default:
                        return null;
                }
                string content = Convert.ToBase64String(stream.GetBuffer());
                Application application = new ApplicationClass();
                application.GetHierarchy(null, HierarchyScope.hsPages, out str2);
                XDocument document = XDocument.Parse(str2);
                XNamespace ns = document.Root.Name.Namespace;
                XElement element = document.Descendants((XName)(ns + "Page")).FirstOrDefault<XElement>();
                string str3 = element.Attribute("ID").Value;
                if (element == null)
                {
                    goto Label_0493;
                }
                string str8 = extension.ToLower().Substring(1);
                if (str8 == null)
                {
                    goto Label_0222;
                }
                if (!(str8 == "jpg"))
                {
                    if (str8 == "png")
                    {
                        goto Label_0210;
                    }
                    if (str8 == "emf")
                    {
                        goto Label_0219;
                    }
                    goto Label_0222;
                }
                string extensionType = "jpg";
                goto Label_022B;
            Label_0210:
                extensionType = "png";
                goto Label_022B;
            Label_0219:
                extensionType = "emf";
                goto Label_022B;
            Label_0222:
                extensionType = "auto";
            Label_022B: ;
                XDocument document2 = new XDocument(new object[] { new XElement((XName)(ns + "Page"), new XElement((XName)(ns + "Outline"), new XElement((XName)(ns + "OEChildren"), new XElement((XName)(ns + "OE"), new XElement((XName)(ns + "Image"), new object[] { new XAttribute("format", extensionType), new XAttribute("originalPageNumber", "0"), new XElement((XName)(ns + "Position"), new object[] { new XAttribute("x", "0"), new XAttribute("y", "0"), new XAttribute("z", "0") }), new XElement((XName)(ns + "Size"), new object[] { new XAttribute("width", bitmap.Width.ToString()), new XAttribute("height", bitmap.Height.ToString()) }), new XElement((XName)(ns + "Data"), content) }))))) });
                document2.Root.SetAttributeValue("ID", str3);
                application.UpdatePageContent(document2.ToString(), DateTime.MinValue);
                application.NavigateTo(str3);
                int num = 10;
                string str5 = "";
                while (num > 0)
                {
                    string str6;
                    Thread.Sleep(0x3e8);
                    application.GetPageContent(str3, out str6, PageInfo.piBinaryDataSelection);
                    if (predicate == null)
                    {
                        predicate = p => (p is XElement) && ((p as XElement).Name == (ns + "OCRText"));
                    }
                    XNode node = XDocument.Parse(str6).DescendantNodes().FirstOrDefault<XNode>(predicate);
                    if (node != null)
                    {
                        return (node as XElement).Value;
                    }
                    num++;
                }
                return str5;
            Label_0493:
                str7 = "";
            }
            return str7;
        }      
        private void LoadOrCreatePage()
        {
            string str;
            this._app.GetHierarchy(string.Empty, HierarchyScope.hsPages, out str);
            XElement element = XDocument.Parse(str).Descendants().FirstOrDefault<XElement>(e => e.Name.LocalName.Equals("Section"));
            if (element == null)
            {
                throw new OcrException("No section found");
            }
            string bstrSectionID = element.Attribute("ID").Value;
            this._app.CreateNewPage(bstrSectionID, out this._pageId);
            this.Reload();
        }

        public string ReadOcrText()
        {
            //var page = this._document.Element(XName.Get("Page", ns));
            //var outline = page.Element(XName.Get("Outline", ns));
            //var OEChildren=outline.Element(XName.Get("OEChildren", ns));
            //var OE = OEChildren.Element(XName.Get("OE", ns));
            //var Image = OE.Element(XName.Get("Image", this.ns));
            //var ocrData = Image.Element(XName.Get("OCRData", this.ns));           

            //if (ocrData == null)
            //    return null;

            //var ocrText = ocrData.Element(XName.Get("OCRText", ns)).Value;
            //return ocrText;
            XElement node = this._document.Element(XName.Get("Page", this.ns)).Element(XName.Get("Outline", this.ns)).Element(XName.Get("OEChildren", this.ns)).Element(XName.Get("OE", this.ns)).Element(XName.Get("Image", this.ns)).Element(XName.Get("OCRData", this.ns));
            if (node != null)
            {
                string text = node.Element(XName.Get("OCRText", this.ns)).Value;
                XmlNamespaceManager resolver = new XmlNamespaceManager(new XmlDocument().NameTable);
                resolver.AddNamespace("one", this._document.Root.Name.NamespaceName);
                IEnumerable<string[]> xml = from o in node.XPathSelectElements("//one:OCRToken", resolver) select new string[] { o.Attribute("line").Value, o.Attribute("x").Value, o.Attribute("width").Value, o.Attribute("y").Value, o.Attribute("height").Value };
                List<EndCharPosition> list = this.LineEdgeCheck(text, xml);
                string[] strArray = new string[list.Count];
                for (int i = 0; (i < strArray.Length) && (i < list.Count); i++)
                {
                    if (list[i].heigh < 50f)
                    {
                        strArray[i] = list[i].Text;
                    }
                    else
                    {
                        strArray[i] = list[i].isedge ? (list[i].Text + "\x00b6") : list[i].Text;
                    }
                }
                return string.Join("\n", strArray);
            }
            return null;
        }
        public List<EndCharPosition> LineEdgeCheck(IEnumerable<string[]> xml)
        {
            EndCharPosition position;
            int num9;
            List<EndCharPosition> list = new List<EndCharPosition>();
            string[] strArray = new string[] { "0", "0", "0" };
            int num = 0;
            foreach (string[] strArray2 in xml)
            {
                if (strArray[0] != strArray2[0])
                {
                    position = new EndCharPosition
                    {
                        line = int.Parse(strArray[0]),
                        right = float.Parse(strArray[1]) + float.Parse(strArray[2]),
                        charcount = num
                    };
                    list.Add(position);
                    num = 0;
                }
                strArray = strArray2;
                num++;
            }
            if (xml.Count<string[]>() > 0)
            {
                num++;
                position = new EndCharPosition
                {
                    line = int.Parse(strArray[0]),
                    right = float.Parse(strArray[1]) + float.Parse(strArray[2]),
                    charcount = num
                };
                list.Add(position);
            }
            list.Sort();
            float num2 = 0f;
            int num3 = (int)Math.Min(Math.Max((double)(list.Count * 0.4), (double)20.0), (double)list.Count);
            int num4 = 0;
            int num5 = 0;
            for (int i = 0; i < num3; i++)
            {
                float right = list[i].right;
                int num8 = 0;
                num9 = i + 1;
                while ((num9 < (20 + i)) && (num9 < list.Count))
                {
                    float num10 = list[num9].right;
                    if ((list[i].charcount > 0x12) && (((right - num10) / right) < 0.03f))
                    {
                        num8++;
                    }
                    else
                    {
                        break;
                    }
                    num9++;
                }
                if (num8 > num4)
                {
                    num4 = num8;
                    num5 = i;
                }
            }
            if ((num4 > 5) || ((list.Count < 10) && (num4 > 1)))
            {
                float num11 = 0f;
                for (num9 = 0; num9 < num4; num9++)
                {
                    num11 += list[num5 + num9].right;
                }
                num2 = num11 / ((float)num4);
                for (num9 = 0; num9 < list.Count; num9++)
                {
                    if (((Math.Abs((float)(num2 - list[num9].right)) / num2) < 0.03f) && (list[num9].charcount > 15))
                    {
                        list[num9].isedge = true;
                    }
                }
            }
            list.Sort(new Comparison<EndCharPosition>(this.CompareEndCharPositionByLine));
            return list;
        }

        public List<EndCharPosition> LineEdgeCheck(string text, IEnumerable<string[]> xml)
        {
            EndCharPosition position;
            int num20;
            EndCharPosition position2;
            if (string.IsNullOrEmpty(text))
            {
                return new List<EndCharPosition>();
            }
            string[] strArray = text.Split(new char[] { '\n' });
            List<EndCharPosition> list = new List<EndCharPosition>();
            string[] strArray2 = new string[] { "0", "0", "0", "0", "0" };
            int num = 0;
            float num2 = 0f;
            float num3 = 0f;
            float num4 = 0f;
            float num5 = 0f;
            float num6 = 0f;
            float num7 = 0f;
            float num8 = 0f;
            float num9 = 0f;
            int num10 = 0;
            foreach (string[] strArray3 in xml)
            {
                if (strArray2[0] != strArray3[0])
                {
                    position = new EndCharPosition();
                   
                        position.line = int.Parse(strArray2[0]);
                        position.Text = strArray[position.line];
                        position.avgcharheight = num8 / num7;
                       position. heigh = float.Parse(strArray2[3]) + float.Parse(strArray2[4]);
                        position.right = float.Parse(strArray2[1]) + float.Parse(strArray2[2]);
                        position.charcount = num;
                   
                    list.Add(position);
                    num3 += num6;
                    num2 += num7;
                    num9 += num8;
                    num10 += num;
                    num = 0;
                    num6 = 0f;
                    num7 = 0f;
                    num8 = 0f;
                }
                float num11 = float.Parse(strArray3[2]);
                float num12 = float.Parse(strArray3[4]);
                num8 += num11 * num12;
                num7 += num11;
                num6 += num12;
                num++;
                strArray2 = strArray3;
            }
            if (strArray2[0] != "0")
            {
                position = new EndCharPosition();
                
                    position.line = int.Parse(strArray2[0]);
                     position.Text = strArray[position.line];
                    position. avgcharheight = num6 / ((float)num);
                     position.heigh = float.Parse(strArray2[3]) + float.Parse(strArray2[4]);
                     position.right = float.Parse(strArray2[1]) + float.Parse(strArray2[2]);
                    position.charcount = num;
                
                list.Add(position);
                num3 += num6;
                num2 += num7;
                num9 += num8;
                num10 += num;
            }
            if (num10 > 0)
            {
                num4 = num9 / num2;
                num5 = num2 / ((float)num10);
            }
            list.Sort();
            float num13 = 0f;
            int num14 = (int)Math.Min(Math.Max((double)(list.Count * 0.4), (double)20.0), (double)list.Count);
            int num15 = 0;
            int num16 = 0;
            int num17 = 0;
            while (num17 < num14)
            {
                float right = list[num17].right;
                int num19 = 0;
                num20 = num17 + 1;
                while ((num20 < (20 + num17)) && (num20 < list.Count))
                {
                    float num21 = list[num20].right;
                    if ((list[num17].charcount > 0x12) && ((right - num21) < (2f * num5)))
                    {
                        num19++;
                    }
                    else
                    {
                        break;
                    }
                    num20++;
                }
                if (num19 > num15)
                {
                    num15 = num19;
                    num16 = num17;
                }
                num17++;
            }
            if ((num15 > 5) || ((list.Count < 10) && (num15 > 1)))
            {
                float num22 = 0f;
                for (num20 = 0; num20 < num15; num20++)
                {
                    num22 += list[num16 + num20].right;
                }
                num13 = num22 / ((float)num15);
            }
            list.Sort(new Comparison<EndCharPosition>(this.CompareEndCharPositionByLine));
            int count = list.Count;
            for (num20 = 1; num20 < count; num20++)
            {
                position2 = list[num20];
                if (position2.right > (num13 + (2f * num5)))
                {
                    position2.isoveredgeline = true;
                    list.RemoveAt(num20);
                    count--;
                }
            }
            for (num20 = 1; num20 < list.Count; num20++)
            {
                position2 = list[num20 - 1];
                EndCharPosition position3 = list[num20];
                if (position3.isoveredgeline)
                {
                    break;
                }
                if (position3.heigh < (position2.heigh + (num4 / 2f)))
                {
                    for (num17 = num20 - 1; num17 >= 0; num17--)
                    {
                        EndCharPosition position4 = list[num17];
                        if (position3.heigh > (position4.heigh - (num4 / 2f)))
                        {
                            position4.right = position3.right;
                            position4.Text = position4.Text + position3.Text;
                            position4.charcount += position3.charcount;
                            list.RemoveAt(num20);
                            num20--;
                            break;
                        }
                    }
                }
            }
            for (num20 = 0; num20 < list.Count; num20++)
            {
                if (Math.Abs((float)(num13 - list[num20].right)) < (2f * num5))
                {
                    list[num20].isedge = true;
                }
            }
            return list;
        }

        public void Reload()
        {
            string str;
            this._app.GetPageContent(this._pageId, out str, PageInfo.piBinaryData);
            this._document = XDocument.Parse(str);
            this.ns = this._document.Root.Name.Namespace.ToString();
            this.onenotedoc = this._document;
        }

        public void Save()
        {
            string bstrPageChangesXmlIn = this._document.ToString();
            this._app.UpdatePageContent(bstrPageChangesXmlIn);
            this._app.NavigateTo(this._pageId);
        }

        private string ToBase64(Image image)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                image.Save(stream, ImageFormat.Png);
                return Convert.ToBase64String(stream.ToArray());
            }
        }

        private XElement Oe
        {
            get
            {
                XElement content = this._document.Root.Element(XName.Get("Outline", "http://schemas.microsoft.com/office/onenote/2013/onenote"));
                if (content == null)
                {
                    content = XElement.Parse("<one:Outline xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\"><one:OEChildren><one:OE><one:T><![CDATA[A]]></one:T></one:OE></one:OEChildren></one:Outline>");
                    this._document.Root.Add(content);
                }
                return content.Element(XName.Get("OEChildren", "http://schemas.microsoft.com/office/onenote/2013/onenote")).Element(XName.Get("OE", "http://schemas.microsoft.com/office/onenote/2013/onenote"));
            }
        }

        public class EndCharPosition : IComparable<OnenotePages.EndCharPosition>
        {
            public float avgcharheight = 0f;
            public float bottom = 0f;
            public int charcount = 0;
            public float heigh = 0f;
            public bool isedge = false;
            public bool isoveredgeline = false;
            public int line;
            public float right;
            public string Text = "";

            public int CompareTo(OnenotePages.EndCharPosition other)
            {
                if (this.right < other.right)
                {
                    return 1;
                }
                if (this.right == other.right)
                {
                    return 0;
                }
                return -1;
            }

            public override string ToString()
            {
                return string.Format("{0},{1},{2},{3},{4},{5},{6}", new object[] { this.line, this.charcount, this.avgcharheight, this.right, this.heigh, this.isedge, this.Text });
            }
        }
    }

}
