using System;
using System.Linq;
using System.Collections.Generic;
using io = System.IO;
using wrd = Aspose.Words;
using excl = Microsoft.Office.Interop.Excel;
using wd = Microsoft.Office.Interop.Word;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace XlsDocImgToPdf
{
    class Program
    {
        static string fails = string.Empty;  //转换失败的文件
        static int failnum = 0;      //转换失败的文件数量
        static string outpath = string.Empty;

        static void Main(string[] args)
        {
            Console.WriteLine("Start.. \n");
            //需要跳过的文件
            var skipfiles = System.Configuration.ConfigurationSettings.AppSettings["skipfiles"];
            var skips = skipfiles.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            //需转换文件目录
            string inpath = System.Configuration.ConfigurationSettings.AppSettings["inpath"];
            //转换成功的文件目录
            outpath = System.Configuration.ConfigurationSettings.AppSettings["outpath"];
            if (!io.Directory.Exists(outpath)) { io.Directory.CreateDirectory(outpath); }
            //循环转换目录下所有文件及子目录内文件
            loop(new io.DirectoryInfo(inpath), null, skips);

            Console.WriteLine(failnum > 0 ? $"存在 {failnum} 个转换失败的文件\n" : "全部转换成功\n");
            if (!string.IsNullOrEmpty(fails))
            {
                Console.WriteLine(fails);
            }
            Console.WriteLine("Please enter any key to close..");
            Console.ReadKey(true);
        }

        /// <summary>
        /// 递归循环目录
        /// </summary>
        /// <param name="dic">目录</param>
        /// <param name="transed">已完成转换的目录</param>
        /// <param name="skips">需要跳过的文件</param>
        /// <param name="fails">转换失败的文件s</param>
        /// <param name="failnum">转换失败的文件个数</param>
        public static void loop(io.DirectoryInfo dic, List<string> transed, string[] skips)
        {
            transfor(dic.FullName, skips);
            if (transed == null) transed = new List<string>();
            transed.Add(dic.FullName);
            io.DirectoryInfo root = new io.DirectoryInfo(dic.FullName);
            io.DirectoryInfo[] dics = root.GetDirectories();
            foreach (io.DirectoryInfo d in dics)
            {
                if (!transed.Contains(d.FullName))
                {
                    loop(d, transed, skips);
                }
            }
        }

        /// <summary>
        /// 转换指定目录内的文件
        /// </summary>
        /// <param name="inpath"></param>
        /// <param name="skips"></param>
        /// <returns></returns>
        public static void transfor(string inpath, string[] skips)
        {
            io.DirectoryInfo root = new io.DirectoryInfo(inpath);
            io.FileInfo[] files = root.GetFiles();
            foreach (io.FileInfo file in files)
            {
                var st = DateTime.Now;
                //outpath = inpath;  //转换成功的文件和原文件放在同一目录

                //check skip
                if (skips != null)
                {
                    if (skips.Contains(file.Name))
                    {
                        Console.WriteLine(file.FullName + " \n跳过，设置为跳过的项\n");
                        continue;
                    }
                }
                if (file.Extension.ToLower() == ".pdf")
                {
                    Console.WriteLine(file.FullName + " \n跳过，pdf文件\n");
                    continue;
                }

                //check fail
                if (file.FullName.Contains("\\fail\\"))
                {
                    Console.WriteLine(file.FullName + " \n跳过，转换失败的项\n");
                    continue;
                }

                //check transed
                var _matchstr = file.FullName.Replace(inpath, outpath);
                _matchstr = _matchstr.Substring(0, _matchstr.IndexOf(".")) + ".pdf";
                if (io.File.Exists(_matchstr))
                {
                    Console.WriteLine(file.FullName + " \n跳过，已经转换的项\n");
                    continue;
                }

                //word to pdf
                if (new string[] { ".doc", ".docx" }.Contains(file.Extension.ToLower()))
                {
                    Console.WriteLine(file.FullName + " 开始转换为PDF");
                    var infile = file.FullName;
                    var outfile = outpath + @"\" + file.Name.Substring(0, file.Name.IndexOf('.')) + ".pdf";
                    if (wordtopdf(infile, outfile))
                    {
                        Console.WriteLine("转换成功");
                        Console.WriteLine("耗时：" + (DateTime.Now - st).ToString() + "\n");
                    }
                    else
                    {
                        fails += file.FullName + "\n"; failnum++;
                        Console.WriteLine("转换失败\n");
                    }
                    continue;
                }

                //execl to pdf
                if (new string[] { ".xls", ".xlsx" }.Contains(file.Extension.ToLower()))
                {
                    Console.WriteLine(file.FullName + " 开始转换为PDF");
                    var infile = file.FullName;
                    var outfile = outpath + @"\" + file.Name.Substring(0, file.Name.IndexOf('.')) + ".pdf";
                    if (execltopdf(infile, outfile))
                    {
                        Console.WriteLine("转换成功");
                        Console.WriteLine("耗时：" + (DateTime.Now - st).ToString() + "\n");
                    }
                    else
                    {
                        fails += file.FullName + "\n"; failnum++;
                        Console.WriteLine("转换失败\n");
                    }
                    continue;
                }

                //imgage to pdf
                if (new string[] { ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff" }.Contains(file.Extension.ToLower()))
                {
                    Console.WriteLine(file.FullName + " 开始转换为PDF");
                    var infile = file.FullName;
                    var outfile = outpath + @"\" + file.Name.Substring(0, file.Name.IndexOf('.')) + ".pdf";
                    if (imgtopdf(infile, outfile))
                    {
                        Console.WriteLine("转换成功");
                        Console.WriteLine("耗时：" + (DateTime.Now - st).ToString() + "\n");
                    }
                    else
                    {
                        fails += file.FullName + "\n"; failnum++;
                        Console.WriteLine("转换失败\n");
                    }
                    continue;
                }

                //other file
                Console.WriteLine(file.FullName + " \n跳过，无需转换\n");
            }
        }

        /// <summary>
        /// word to pdf
        /// </summary>
        /// <param name="infile"></param>
        /// <param name="outfile"></param>
        /// <returns></returns>
        public static bool wordtopdf(string infile, string outfile)
        {
            wd.Application application = null;
            wd.Document document = null;
            try
            {
                application = new wd.Application();
                application.Visible = false;
                document = application.Documents.Open(infile);
                document.ExportAsFixedFormat(outfile, wd.WdExportFormat.wdExportFormatPDF);
                return true;
            }
            catch (Exception ex)
            {
                copyto(infile);
                return false;
            }
            finally
            {
                if (document != null)
                {
                    document.Close();
                    document = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
            }
        }

        /// <summary>
        /// word to pdf（此方法文件过大会有pdf内容不全的问题）
        /// ref：http://blog.51cto.com/xiaoshuaigege/1889700
        /// </summary>
        /// <param name="infile"></param>
        /// <param name="outfile"></param>
        /// <returns></returns>
        public static bool wordtopdf2(string infile, string outfile)
        {
            try
            {
                wrd.Document doc = new wrd.Document(infile);
                doc.Save(outfile, wrd.SaveFormat.Pdf);
                return true;
            }
            catch (Exception ex)
            {
                copyto(infile);
                return false;
            }
        }

        /// <summary>
        /// execl to pdf
        /// 需要安装office 2007 还有一个office2007的插件OfficeSaveAsPDFandXPS
        /// 下载地址 http://www.microsoft.com/downloads/details.aspx?FamilyId=4D951911-3E7E-4AE6-B059-A2E79ED87041&displaylang=en
        /// </summary>
        /// <param name="infile"></param>
        /// <param name="outfile"></param>
        /// <returns></returns>
        public static bool execltopdf(string infile, string outfile)
        {
            excl.Application application = null;
            excl.Workbook workBook = null;
            try
            {
                //object _missing = Type.Missing;
                object _missing = System.Reflection.Missing.Value;
                application = new excl.Application();
                application.Visible = true;
                workBook = application.Workbooks.Open(infile, true, false, _missing, _missing, _missing, true, _missing,
                    _missing, _missing, _missing, _missing, false, _missing, _missing);

                //指定execl worksheet宽度为1页
                excl.Worksheet worksheet = workBook.Worksheets[1];
                worksheet.PageSetup.Zoom = false;
                worksheet.PageSetup.FitToPagesWide = 1;
                worksheet.PageSetup.FitToPagesTall = false;

                workBook.ExportAsFixedFormat(excl.XlFixedFormatType.xlTypePDF, outfile, excl.XlFixedFormatQuality.xlQualityStandard, true, false, _missing, _missing, _missing, _missing);
                return true;
            }
            catch (Exception ex)
            {
                copyto(infile);
                return false;
            }
            finally
            {
                if (workBook != null)
                {
                    workBook.Close(false, Type.Missing, Type.Missing);
                    workBook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
            }
        }

        /// <summary>
        /// image to pdf
        /// </summary>
        /// <param name="infile"></param>
        /// <param name="outfile"></param>
        /// <returns></returns>
        public static bool imgtopdf(string infile, string outfile)
        {
            Document document = null;
            try
            {
                document = new Document(PageSize.A4, 25, 25, 25, 25);
                using (var stream = new io.FileStream(outfile, io.FileMode.Create, io.FileAccess.Write, io.FileShare.None))
                {
                    PdfWriter.GetInstance(document, stream);
                    document.Open();
                    using (var imageStream = new io.FileStream(infile, io.FileMode.Open, io.FileAccess.Read, io.FileShare.ReadWrite))
                    {
                        var image = Image.GetInstance(imageStream);
                        if (image.Height > PageSize.A4.Height - 25)
                        {
                            image.ScaleToFit(PageSize.A4.Width - 25, PageSize.A4.Height - 25);
                        }
                        else if (image.Width > PageSize.A4.Width - 25)
                        {
                            image.ScaleToFit(PageSize.A4.Width - 25, PageSize.A4.Height - 25);
                        }
                        image.Alignment = Image.ALIGN_MIDDLE;
                        //document.NewPage();
                        document.Add(image);
                    }
                    document.Close();
                    return true;
                }
            }
            catch (Exception ex)
            {
                copyto(infile);
                return false;
            }
        }

        /// <summary>
        /// 拷贝转换失败的文件到其他目录
        /// </summary>
        /// <param name="infile"></param>
        public static void copyto(string infile)
        {
            try
            {
                var failpath = System.Configuration.ConfigurationSettings.AppSettings["failpath"];
                if (!io.Directory.Exists(failpath))
                {
                    io.Directory.CreateDirectory(failpath);
                }
                var copyto = failpath + "\\" + infile.Substring(infile.LastIndexOf('\\') + 1);
                io.File.Copy(infile, copyto, true);
            }
            catch { }
        }
    }
}
