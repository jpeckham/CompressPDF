using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Image = iTextSharp.text.Image;
using Rectangle = iTextSharp.text.Rectangle;

namespace CompressPDF
{
    class Program
    {
        static void Main(string[] args)
        {
            Document document = new Document(PageSize.LETTER, 0f, 0f, 0f, 0f);
            string inputFile = "Heroes Unlimited RPG - 2E.pdf";
            string destinationFile = "OUT" + inputFile;
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(destinationFile, FileMode.Create));
            writer.SetFullCompression();
            PdfDate st = new PdfDate(DateTime.Today);

            document.Open();
            PdfContentByte cb = writer.DirectContent;

            PdfImportedPage page;

            int rotation;

            int iPageNo = 1;

            PdfReader reader = new PdfReader(inputFile);
            int n = reader.NumberOfPages;

            int i = 1;

            while (i <= n)
            {
                document.NewPage();
               
                PdfDictionary pg = reader.GetPageN(i);

                // recursively search pages, forms and groups for images.
                PdfObject obj = FindImageInPDFDictionary(pg);
                if (obj != null)
                {

                    int XrefIndex =
                        Convert.ToInt32(((PRIndirectReference) obj).Number.ToString(System.Globalization.CultureInfo
                            .InvariantCulture));
                    PdfObject pdfObj = reader.GetPdfObject(XrefIndex);
                    PdfStream pdfStrem = (PdfStream) pdfObj;
                    byte[] bytes = PdfReader.GetStreamBytesRaw((PRStream) pdfStrem);
                    if ((bytes != null))
                    {
                        using (System.IO.MemoryStream memStream = new System.IO.MemoryStream(bytes))
                        {
                            memStream.Position = 0;
                            System.Drawing.Image img =i==1 || i==n ? System.Drawing.Image.FromStream(memStream)://let the first and last page stay the same size.
                                BitmapToGrayscale(ImageHelper.ResizeImage(System.Drawing.Image.FromStream(memStream), .35M));
                            
                           
                                ImageFormat format = img.PixelFormat == PixelFormat.Format1bppIndexed
                                                     || img.PixelFormat == PixelFormat.Format4bppIndexed
                                                     || img.PixelFormat == PixelFormat.Format8bppIndexed
                                    ? ImageFormat.Tiff
                                    : ImageFormat.Jpeg;
                                
                                var pdfImage = iTextSharp.text.Image.GetInstance(img, format);
                                pdfImage.Alignment = Element.ALIGN_CENTER;
                                pdfImage.ScaleToFit(document.PageSize.Width - 10, document.PageSize.Height - 10);
                            pdfImage.ColorTransform = 1;
                                //page = writer.GetImportedPage(reader, i);
                            document.Add(pdfImage);

                        }
                    }
                }
                
                i++;
                Console.SetCursorPosition(0,0);
                Console.Write($"{((i*100)/(n)).ToString().PadLeft(3)}%");
            }
            Console.WriteLine("Complete");
            document.Close();

           // ExtractImagesFromPDF(inputFile, ".\\");
        }

        public static Bitmap BitmapToGrayscale4bpp(Bitmap source)
        {
            // Create target image.
            int width = source.Width;
            int height = source.Height;
            Bitmap target = new Bitmap(width, height, PixelFormat.Format4bppIndexed);
            // Set the palette to discrete shades of gray
            ColorPalette palette = target.Palette;
            for (int i = 0; i < palette.Entries.Length; i++)
            {
                int cval = 17 * i;
                palette.Entries[i] = Color.FromArgb(0, cval, cval, cval);
            }
            target.Palette = palette;

            // Lock bits so we have direct access to bitmap data
            BitmapData targetData = target.LockBits(new System.Drawing.Rectangle(0, 0, width, height),
                ImageLockMode.ReadWrite, PixelFormat.Format4bppIndexed);
            BitmapData sourceData = source.LockBits(new System.Drawing.Rectangle(0, 0, width, height),
                ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

            unsafe
            {
                for (int r = 0; r < height; r++)
                {
                    byte* pTarget = (byte*)(targetData.Scan0 + r * targetData.Stride);
                    byte* pSource = (byte*)(sourceData.Scan0 + r * sourceData.Stride);
                    byte prevValue = 0;
                    for (int c = 0; c < width; c++)
                    {
                        byte colorIndex = (byte)((((*pSource) * 0.3 + *(pSource + 1) * 0.59 + *(pSource + 2) * 0.11)) / 16);
                        if (c % 2 == 0)
                            prevValue = colorIndex;
                        else
                            *(pTarget++) = (byte)(prevValue | colorIndex << 4);

                        pSource += 3;
                    }
                }
            }

            target.UnlockBits(targetData);
            source.UnlockBits(sourceData);
            return target;
        }
        public static void ExtractImagesFromPDF(string sourcePdf, string outputPath)
        {
            // NOTE:  This will only get the first image it finds per page.
            PdfReader pdf = new PdfReader(sourcePdf);
            //RandomAccessFileOrArray raf = new iTextSharp.text.pdf.RandomAccessFileOrArray(sourcePdf);

            try
            {
                for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++)
                {
                    PdfDictionary pg = pdf.GetPageN(pageNumber);

                    // recursively search pages, forms and groups for images.
                    PdfObject obj = FindImageInPDFDictionary(pg);
                    if (obj != null)
                    {

                        int XrefIndex = Convert.ToInt32(((PRIndirectReference)obj).Number.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        PdfObject pdfObj = pdf.GetPdfObject(XrefIndex);
                        PdfStream pdfStrem = (PdfStream)pdfObj;
                        byte[] bytes = PdfReader.GetStreamBytesRaw((PRStream)pdfStrem);
                        if ((bytes != null))
                        {
                            using (System.IO.MemoryStream memStream = new System.IO.MemoryStream(bytes))
                            {
                                memStream.Position = 0;
                                System.Drawing.Image img = ImageHelper.ResizeImage(System.Drawing.Image.FromStream(memStream),.33M);
                                // must save the file while stream is open.
                                if (!Directory.Exists(outputPath))
                                    Directory.CreateDirectory(outputPath);

                                string path = Path.Combine(outputPath, String.Format(@"{0}.jpg", pageNumber));
                                System.Drawing.Imaging.EncoderParameters parms = new System.Drawing.Imaging.EncoderParameters(1);
                                parms.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Compression, 0);
                                var encoders = ImageCodecInfo.GetImageEncoders();
                                System.Drawing.Imaging.ImageCodecInfo jpegEncoder = encoders.FirstOrDefault(p => p.CodecName == "Built-in JPEG Codec");
                                img.Save(path, jpegEncoder, parms);
                            }
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                pdf.Close();
                //raf.Close();
            }


        }
        public static Bitmap BitmapToGrayscale(Bitmap source)
        {
            // Create target image.
            int width = source.Width;
            int height = source.Height;
            Bitmap target = new Bitmap(width, height, PixelFormat.Format8bppIndexed);
            // Set the palette to discrete shades of gray
            ColorPalette palette = target.Palette;
            for (int i = 0; i < palette.Entries.Length; i++)
            {
                palette.Entries[i] = Color.FromArgb(0, i, i, i);
            }
            target.Palette = palette;

            // Lock bits so we have direct access to bitmap data
            BitmapData targetData = target.LockBits(new System.Drawing.Rectangle(0, 0, width, height),
                ImageLockMode.ReadWrite, PixelFormat.Format8bppIndexed);
            BitmapData sourceData = source.LockBits(new System.Drawing.Rectangle(0, 0, width, height),
                ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

            unsafe
            {
                for (int r = 0; r < height; r++)
                {
                    byte* pTarget = (byte*)(targetData.Scan0 + r * targetData.Stride);
                    byte* pSource = (byte*)(sourceData.Scan0 + r * sourceData.Stride);
                    for (int c = 0; c < width; c++)
                    {
                        byte colorIndex = (byte)(((*pSource) * 0.3 + *(pSource + 1) * 0.59 + *(pSource + 2) * 0.11));
                        *pTarget = colorIndex;
                        pTarget++;
                        pSource += 3;
                    }
                }
            }

            target.UnlockBits(targetData);
            source.UnlockBits(sourceData);
            return target;
        }
        private static PdfObject FindImageInPDFDictionary(PdfDictionary pg)
        {
            PdfDictionary res =
                (PdfDictionary)PdfReader.GetPdfObject(pg.Get(PdfName.RESOURCES));


            PdfDictionary xobj =
                (PdfDictionary)PdfReader.GetPdfObject(res.Get(PdfName.XOBJECT));
            if (xobj != null)
            {
                foreach (PdfName name in xobj.Keys)
                {

                    PdfObject obj = xobj.Get(name);
                    if (obj.IsIndirect())
                    {
                        PdfDictionary tg = (PdfDictionary)PdfReader.GetPdfObject(obj);

                        PdfName type =
                            (PdfName)PdfReader.GetPdfObject(tg.Get(PdfName.SUBTYPE));

                        //image at the root of the pdf
                        if (PdfName.IMAGE.Equals(type))
                        {
                            return obj;
                        }// image inside a form
                        else if (PdfName.FORM.Equals(type))
                        {
                            return FindImageInPDFDictionary(tg);
                        } //image inside a group
                        else if (PdfName.GROUP.Equals(type))
                        {
                            return FindImageInPDFDictionary(tg);
                        }

                    }
                }
            }

            return null;

        }

        public static class ImageHelper
        {
            /// <summary>
            /// Resize the image to the specified width and height.
            /// </summary>
            /// <param name="image">The image to resize.</param>
            /// <param name="width">The width to resize to.</param>
            /// <param name="height">The height to resize to.</param>
            /// <returns>The resized image.</returns>
            public static Bitmap ResizeImage(System.Drawing.Image image, int width, int height)
            {
                var destRect = new System.Drawing.Rectangle(0, 0, width, height);
                var destImage = new System.Drawing.Bitmap(width, height);

                destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

                using (var graphics = Graphics.FromImage(destImage))
                {
                    graphics.CompositingMode = CompositingMode.SourceCopy;
                    graphics.CompositingQuality = CompositingQuality.HighQuality;
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.SmoothingMode = SmoothingMode.HighQuality;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                    using (var wrapMode = new ImageAttributes())
                    {
                        wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                        graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                    }
                }

                return destImage;
            }

            public static Bitmap ResizeImage(System.Drawing.Image image, decimal percentage)
            {
                int width = (int)Math.Round(image.Width * percentage, MidpointRounding.AwayFromZero);
                int height = (int)Math.Round(image.Height * percentage, MidpointRounding.AwayFromZero);
                return ResizeImage(image, width, height);
            }
        }
    }
}
