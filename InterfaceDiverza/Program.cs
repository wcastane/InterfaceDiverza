using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;

using System.Xml;
using System.Xml.Linq;

namespace InterfaceDiverza
{
    class Program
    {
        static void Main(string[] args)
        {
            
            string inputpath = InterfaceDiverza.Properties.Settings.Default.inputpath.ToString();
            string outputpath = InterfaceDiverza.Properties.Settings.Default.outputpath.ToString();
            ArrangeFiles(inputpath, outputpath);
        }

        private static void ArrangeFiles(string inputpath, string outputpath)
        {

            //Consultamos el listado de archivos sobre el directorio.
            //Si existen archivos .zip son descomprimidos en el mismo folder en una carpeta con el mismo nombre del archivo.
            //Si es un archivo .xml se consulta la información para acomodarlos dependiendo de los criterios
            string[] files = Directory.GetFiles(inputpath);
            foreach (var f in files)
            {
                string ext = Path.GetExtension(f).ToLower();
                string localpath = string.Empty;

                switch (ext)
                {
                    case ".zip":
                        localpath = Path.Combine(Path.GetDirectoryName(f), Path.GetFileNameWithoutExtension(f));
                        ExtractZipFile(f, string.Empty, localpath);
                        File.Delete(f);
                        Console.WriteLine($"Extracted {f}");
                        break;

                    case ".xml":
                        string emisor = string.Empty;
                        string receptor = string.Empty;
                        string fecha = string.Empty;

                        XDocument xml = XDocument.Load(f);
                        try
                        {
                            XNamespace cfdi = "http://www.sat.gob.mx/cfd/3";
                            emisor = (from item in xml.Descendants(cfdi + "Emisor") select item.Attribute("rfc").Value).First();
                            receptor = (from item in xml.Descendants(cfdi + "Receptor") select item.Attribute("rfc").Value).First();
                            fecha = (from item in xml.Descendants(cfdi + "Comprobante") select item.Attribute("fecha").Value).First();
                        }
                        catch (InvalidOperationException ex)
                        {
                            XNamespace cfdi = "http://www.sat.gob.mx/cfd/2";
                            emisor = (from item in xml.Descendants(cfdi + "Emisor") select item.Attribute("rfc").Value).First();
                            receptor = (from item in xml.Descendants(cfdi + "Receptor") select item.Attribute("rfc").Value).First();
                            fecha = (from item in xml.Descendants(cfdi + "Comprobante") select item.Attribute("fecha").Value).First();
                        }
                        
                        DateTime fecha2 = DateTime.Parse(fecha);

                        if (emisor == "CEM110616B59")
                        {
                            localpath = Path.Combine(outputpath, "clientes", fecha2.Year.ToString(), fecha2.Month.ToString());
                            if (!Directory.Exists(localpath))
                                Directory.CreateDirectory(localpath);
                            File.Copy(f, Path.Combine(localpath, Path.GetFileName(f)),true);
                        }
                        if (receptor == "CEM110616B59")
                        {
                            localpath = Path.Combine(outputpath, "proveedores", fecha2.Year.ToString(), fecha2.Month.ToString());
                            if (!Directory.Exists(localpath))
                                Directory.CreateDirectory(localpath);
                            File.Copy(f, Path.Combine(localpath, Path.GetFileName(f)),true);
                        }
                        if (emisor == "CES110616F46")
                        {
                            localpath = Path.Combine(outputpath, "nomina", fecha2.Year.ToString(), fecha2.Month.ToString());
                            if (!Directory.Exists(localpath))
                                Directory.CreateDirectory(localpath);
                            File.Copy(f, Path.Combine(localpath, Path.GetFileName(f)),true);
                        }
                        if (receptor == "CES110616F46")
                        {
                            localpath = Path.Combine(outputpath, "viajes", fecha2.Year.ToString(), fecha2.Month.ToString());
                            if (!Directory.Exists(localpath))
                                Directory.CreateDirectory(localpath);
                            File.Copy(f, Path.Combine(localpath, Path.GetFileName(f)),true);
                        }
                        File.Delete(f);
                        Console.WriteLine($"Arranged {f}");
                        break;

                    default:
                        break;
                }
            }

            //Consultamos si existen directorios y corremos ArrangeFiles para cada directorio
            string[] directories = Directory.GetDirectories(inputpath);
            foreach (var dir in directories)
            {
                ArrangeFiles(dir,outputpath);
            }

            if (Directory.GetDirectories(inputpath).Count() == 0 && Directory.GetFiles(inputpath).Count() == 0)
            {
                Directory.Delete(inputpath);
                Console.WriteLine($"Se borro {inputpath}");
            }
        }

        private static void ExtractZipFile(string archiveFilenameIn, string password, string outFolder)
        {
            ZipFile zf = null;
            try
            {
                FileStream fs = File.OpenRead(archiveFilenameIn);
                zf = new ZipFile(fs);
                if (!String.IsNullOrEmpty(password))
                {
                    zf.Password = password;     // AES encrypted entries are handled automatically
                }
                foreach (ZipEntry zipEntry in zf)
                {
                    if (!zipEntry.IsFile)
                    {
                        continue;           // Ignore directories
                    }
                    String entryFileName = zipEntry.Name;
                    // to remove the folder from the entry:- entryFileName = Path.GetFileName(entryFileName);
                    // Optionally match entrynames against a selection list here to skip as desired.
                    // The unpacked length is available in the zipEntry.Size property.

                    byte[] buffer = new byte[4096];     // 4K is optimum
                    Stream zipStream = zf.GetInputStream(zipEntry);

                    // Manipulate the output filename here as desired.
                    String fullZipToPath = Path.Combine(outFolder, entryFileName);
                    string directoryName = Path.GetDirectoryName(fullZipToPath);
                    if (directoryName.Length > 0)
                        Directory.CreateDirectory(directoryName);

                    // Unzip file in buffered chunks. This is just as fast as unpacking to a buffer the full size
                    // of the file, but does not waste memory.
                    // The "using" will close the stream even if an exception occurs.
                    using (FileStream streamWriter = File.Create(fullZipToPath))
                    {
                        StreamUtils.Copy(zipStream, streamWriter, buffer);
                    }
                }
            }
            finally
            {
                if (zf != null)
                {
                    zf.IsStreamOwner = true; // Makes close also shut the underlying stream
                    zf.Close(); // Ensure we release resources
                }
            }
        }
    }
}
