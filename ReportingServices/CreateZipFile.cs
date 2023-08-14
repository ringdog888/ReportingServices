using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ICSharpCode.SharpZipLib.Zip;
using System.IO;
using System.Net;

/// <summary>
/// ZipTool 的摘要描述
/// </summary>
public class ZipTool
{

    public void CreateZipFile(List<string> FileList, string ZipPath)
    {
        try
        {
            // Depending on the directory this could be very large and would require more attention
            // in a commercial package.
            
            // 'using' statements guarantee the stream is closed properly which is a big source
            // of problems otherwise.  Its exception safe as well which is great.
            using (ZipOutputStream s = new ZipOutputStream(File.Create(ZipPath)))
            {
                s.SetLevel(9); // 0 - store only to 9 - means best compression

                byte[] buffer = new byte[4096];

                foreach (string file in FileList)
                {
                    WebRequest req = HttpWebRequest.Create(file);

                    using (WebResponse response = req.GetResponse())
                    {
                        string disposition = response.Headers["content-disposition"];
                        byte[] bt = new byte[disposition.Length];
                        for (int i = 0; i < disposition.Length; ++i)
                        {
                            bt[i] = (byte)disposition[i];
                        }
                        string Filename = System.Text.Encoding.GetEncoding("utf-8").GetString(bt).Split('=')[1].Split('.')[0];
                        string FilenameExtension = System.Text.Encoding.GetEncoding("utf-8").GetString(bt).Split('=')[1].Split('.')[1];

                        // Using GetFileName makes the result compatible with XP
                        // as the resulting path is not absolute.
                        ZipEntry entry = new ZipEntry(Path.GetFileName(Filename + Guid.NewGuid() + ".") + FilenameExtension);

                        // Setup the entry data as required.

                        // Crc and size are handled by the library for seakable streams
                        // so no need to do them here.

                        // Could also use the last write time or similar for the file.
                        entry.DateTime = DateTime.Now;
                        s.PutNextEntry(entry);

                        //using (FileStream fs = File.OpenRead(file))
                        using (Stream fs = response.GetResponseStream())
                        {
                            // Using a fixed size buffer here makes no noticeable difference for output
                            // but keeps a lid on memory usage.
                            int sourceBytes;
                            do
                            {
                                sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                s.Write(buffer, 0, sourceBytes);
                            } while (sourceBytes > 0);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Exception during processing {0}", ex);
            // No need to rethrow the exception as for our purposes its handled.
        }
	}
}