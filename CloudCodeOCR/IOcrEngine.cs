using System;
using System.Drawing;
using System.Linq;

namespace CloudCodeOCR
{
    public interface IOcrEngine
    {
        string Recognize(Image image);
    }
}
