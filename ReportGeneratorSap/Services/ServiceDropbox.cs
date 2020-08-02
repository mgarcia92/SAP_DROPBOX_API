using Dropbox.Api;
using Dropbox.Api.Files;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ReportGeneratorSap.Services
{
    public class ServiceDropbox
    {
        public string token;
        public string secretKey;
        public string key;
        public string path;
        public string folder;
        public static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        // private string path_dropbox = "/Plumrose/Info Indicadores/Reportes para la parametrización de SIGC/2020/pruabedecarpeta";

        public ServiceDropbox()
        {
            token = ConfigurationManager.AppSettings["token"].ToString();
            secretKey = ConfigurationManager.AppSettings["secretKey"].ToString();
            key = ConfigurationManager.AppSettings["key"].ToString();
            folder = $"{ConfigurationManager.AppSettings["folder"]}{getNameFolder()}";
            path = ConfigurationManager.AppSettings["path"];
        }

        private DropboxClient getClient()
        {
            return new DropboxClient(token);
        }

        public async Task processUpload()
        {
            try
            {
                using (var client = getClient())
                {
                    //string path = "/Plumrose/Info Indicadores/Reportes para la parametrización de SIGC/2020/pruebaFebrero";
                 
                    var exist = folderExistsDropbox(client, folder).Result;
                    if(!exist) CreateFolder(client, folder);

                    var files = Directory.GetFiles(path);
                    var count = 0;
                    foreach (var file in files)
                    {
                        var info = new FileInfo(file); // info file
                        if (info.Exists && info.Extension.ToLower().Trim() == ".xlsx")
                        {
                            Console.WriteLine("The file is uploading...");
                            var uploaded = await uploadFile(client, info);
                            if (!string.IsNullOrEmpty(uploaded))
                            {
                                count++;
                                string dateFile = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                                logger.Trace($" {dateFile} Uploaded File... with ID =>  + {uploaded}");
                                Console.WriteLine("Uploaded File... with ID => "+uploaded);
                                File.Delete(info.FullName);
                            }
                        }
                    }
                    string date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    Console.WriteLine($"{date} {count} files were uploaded...!");
                    logger.Trace($"{date} {count} files were uploaded...!");
                    
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public string getNameFolder()
        {
            var startWeek = DateTime.Now.AddDays(DayOfWeek.Monday - DateTime.Now.DayOfWeek);
            var MonthName = startWeek.ToString("MMMM", CultureInfo.CreateSpecificCulture("es"));
            var today = startWeek.AddDays(-7).ToString("dd",
                              new CultureInfo("es-VE"));
            DateTime finalDate = startWeek.AddDays(-1);
            var finalDay = finalDate.ToString("dd",
                             new CultureInfo("es-VE"));
            return $"Sem-{today}Al{finalDay}-{MonthName}";
        }

        private async Task<string> uploadFile(DropboxClient client,FileInfo file)
        {
            using (var memoryFile = new MemoryStream(File.ReadAllBytes(file.FullName)))
            {
               var updated =  await client.Files.UploadAsync($"{folder}/{file.Name}",WriteMode.Overwrite.Instance,body: memoryFile);
               var tx = client.Sharing.CreateSharedLinkWithSettingsAsync(folder + "/" + file.Name);
               if (updated.IsFile)
                {
                    return updated.Id;
                }
                else
                {
                    return string.Empty;
                }
            }
        }

        private string convertXlsToXlsx(FileInfo file)
        {
            Console.WriteLine("Compressing Files.....!!");
            var app = new Microsoft.Office.Interop.Excel.Application();
            var xlsFile = file.FullName;
            var wb = app.Workbooks.Open(xlsFile);
            var xlsxFile = xlsFile + "x";
            wb.SaveAs(Filename: xlsxFile, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();
            return xlsxFile;
        }

        public string convertxlsToXSLXV2(string filePath)
        {
            FileInfo file = new FileInfo(filePath);
            Workbook workbook = new Workbook();
         
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                Console.WriteLine("Compressing Files....!!");
                workbook.LoadFromXml(stream);
                string name = Path.GetFileNameWithoutExtension(file.FullName);
                string pathSave = $"{file.DirectoryName}\\{name}.xlsx";
                workbook.SaveToFile(pathSave, ExcelVersion.Version2013);
                return pathSave;
            }
          
        }

        public bool convertoAllToXlsx()
        {
            var files = Directory.GetFiles(path);
           if(files.Length == 0)
            {
                Console.WriteLine("not files in the folder");
                return false;
            }
            foreach (var file in files)
            {
                string date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                FileInfo info = new FileInfo(file);
                if(info.Extension.ToLower().Trim() == ".xls")
                {
                    var result  = convertXlsToXlsx(info);
                    if (!string.IsNullOrEmpty(result))
                    {
                        File.Delete(info.FullName);
                        Console.WriteLine($"{date} File {info.Name} is converted successfully");
                        logger.Trace($"{date} File {info.Name} is converted successfully");                
                    }
                }
                else
                {
                    continue;
                }
            }
            return true;
        }

        public bool CreateFolder(DropboxClient client, string path)
        {
            try
            {         
                var folderArg = new CreateFolderArg(path);
                var folder = client.Files.CreateFolderV2Async(folderArg);
                var result = folder.Result;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool FolderExists(DropboxClient client, string path)
        {
            try
            {         
                var folders = client.Files.ListFolderAsync(path);
                var result = folders.Result;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public bool Delete(DropboxClient client, string path)
        {
            try
            {
                var folders = client.Files.DeleteV2Async(path);
                var result = folders.Result;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public bool Upload(DropboxClient client, string UploadfolderPath, string UploadfileName, string SourceFilePath)
        {
            try
            {
                using (var stream = new MemoryStream(File.ReadAllBytes(SourceFilePath)))
                {
                    var response = client.Files.UploadAsync(UploadfolderPath + "/" + UploadfileName, WriteMode.Overwrite.Instance, body: stream);
                    var rest = response.Result; //Added to wait for the result from Async method  
                }

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }

        public bool Download(DropboxClient client, string DropboxFolderPath, string DropboxFileName, string DownloadFolderPath, string DownloadFileName)
        {
            try
            {
                var response = client.Files.DownloadAsync(DropboxFolderPath + "/" + DropboxFileName);
                var result = response.Result.GetContentAsStreamAsync(); //Added to wait for the result from Async method  

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }

        //public bool CanAuthenticate()
        //{
        //    try
        //    {
        //        if (AppKey == null)
        //        {
        //            throw new ArgumentNullException("AppKey");
        //        }
        //        if (AppSecret == null)
        //        {
        //            throw new ArgumentNullException("AppSecret");
        //        }
        //        return true;
        //    }
        //    catch (Exception)
        //    {
        //        throw;
        //    }

        //}

        public async Task<bool> folderExistsDropbox(DropboxClient client, string path)
        {
            try
            {
                await client.Files.GetMetadataAsync(path);
                return true;
            }
            catch (ApiException<Dropbox.Api.Files.GetMetadataError> e)
            {
                if (e.ErrorResponse.IsPath && e.ErrorResponse.AsPath.Value.IsNotFound)
                {
                  //  Console.WriteLine("Nothing found at path.");
                    return false;
                }
                else
                {
                    // different issue; handle as desired
                   // Console.WriteLine(e);
                    return true;
                }
            }
        }

      


    }
}
