using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace UploadBaoCao
{
    class clsDongBo
    {
        /// <summary>
        ///     ''' 
        ///     ''' </summary>
        ///     ''' <param name="folder_name"></param>
        ///     ''' <param name="username"></param>
        ///     ''' <param name="password"></param>
        ///     ''' <returns></returns>
        public static bool FtpFolderCreate(string ftpAccess, string folder_name, string username, string password)
        {
            System.Net.FtpWebRequest request = (System.Net.FtpWebRequest)System.Net.FtpWebRequest.Create(ftpAccess + folder_name);
            request.Credentials = new System.Net.NetworkCredential(username, password);
            request.Method = System.Net.WebRequestMethods.Ftp.MakeDirectory;

            try
            {
                using (System.Net.FtpWebResponse response = (System.Net.FtpWebResponse)request.GetResponse())
                {
                }
            }
            catch (System.Net.WebException ex)
            {
                System.Net.FtpWebResponse response = (System.Net.FtpWebResponse)ex.Response;
                // an error occurred
                if (response.StatusCode == System.Net.FtpStatusCode.ActionNotTakenFileUnavailable)
                    return false;
            }
            return true;
        }
        /// <summary>
        ///     ''' 
        ///     ''' </summary>
        ///     ''' <param name="ftpAddress">this is the remote server address in the format FTP://server/foldername</param>
        ///     ''' <param name="ftpUser">this is the user name with access to the FTP server</param>
        ///     ''' <param name="ftpPassword">password for the ftpUser</param>
        ///     ''' <returns></returns>
        public static List<string> ListRemoteFiles(string ftpAddress, string folder_name, string ftpUser, string ftpPassword)
        {
            List<string> ListOfFilesOnFTPSite = new List<string>();
            System.Net.FtpWebRequest ftpRequest = null;
            System.Net.FtpWebResponse ftpResponse = null;
            System.IO.StreamReader strReader = null;
            string sline = "";

            try
            {
                ftpRequest = (System.Net.FtpWebRequest)System.Net.WebRequest.Create(ftpAddress);

                {
                    var withBlock = ftpRequest;
                    withBlock.Credentials = new System.Net.NetworkCredential(ftpUser, ftpPassword);
                    withBlock.Method = System.Net.WebRequestMethods.Ftp.ListDirectory;
                }

                ftpResponse = (System.Net.FtpWebResponse)ftpRequest.GetResponse();

                strReader = new System.IO.StreamReader(ftpResponse.GetResponseStream());

                if (strReader != null)
                    sline = strReader.ReadLine();

                while (sline != null)
                {
                    ListOfFilesOnFTPSite.Add(sline.Replace(folder_name, ""));
                    sline = strReader.ReadLine();
                }
            }
            catch (System.Net.WebException ex)
            {
            }

            finally
            {
                if (ftpResponse != null)
                {
                    ftpResponse.Close();
                    ftpResponse = null;
                }

                if (strReader != null)
                {
                    strReader.Close();
                    strReader = null;
                }
            }
            return ListOfFilesOnFTPSite;
            //ListRemoteFiles = ListOfFilesOnFTPSite;

            //ListOfFilesOnFTPSite = null;
        }
        /// <summary>
        ///     ''' 
        ///     ''' </summary>
        ///     ''' <param name="ftpAccess"></param>
        ///     ''' <param name="FilePath"></param>
        ///     ''' <param name="username"></param>
        ///     ''' <param name="password"></param>
        public static bool UpLoadFile(string ftpAccess, string FilePath, string username, string password)
        {
            try
            {
                string[] spl = FilePath.Split('\\');
                ftpAccess += "/" + Convert.ToString(spl[spl.Count() - 1]);
                // Create req
                System.Net.FtpWebRequest mReq = (System.Net.FtpWebRequest)System.Net.WebRequest.Create(ftpAccess);
                // Update properties
                mReq.Credentials = new System.Net.NetworkCredential(username, password);
                mReq.Method = System.Net.WebRequestMethods.Ftp.UploadFile;

                // Read File
                byte[] Mfile = System.IO.File.ReadAllBytes(FilePath);

                // Upload
                System.IO.Stream mStream = mReq.GetRequestStream();
                mStream.Write(Mfile, 0, Mfile.Length);

                // Cleanup
                mStream.Close();
                mStream.Dispose();

                return true;
            }
            catch (Exception ex)
            {
                return false;

            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string DeleteFile(string ftpAddress, string fileName, string ftpUser, string ftpPassword)
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(ftpAddress + "/" + fileName);
            request.Method = WebRequestMethods.Ftp.DeleteFile;
            request.Credentials = new NetworkCredential(ftpUser, ftpPassword);

            using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
            {
                return response.StatusDescription;
            }
        }
        /// <summary>
        ///     ''' 
        ///     ''' </summary>
        ///     ''' <param name="ftpAddress"> this is the remote server address in the format FTP://server/foldername </param>
        ///     ''' <param name="ftpUser">this is the user name with access to the FTP server</param>
        ///     ''' <param name="ftpPassword">password for the ftpUser</param>
        ///     ''' <param name="fileToDownload"> file name you want to download from the FTP folder</param>
        ///     ''' <param name="downloadTargetFolder">local folder to download file to</param>
        ///     ''' <param name="deleteAfterDownload">boolean value to set if you would like to delete the file on the FTP site after download</param>
        ///     ''' <param name="ExceptionInfo"> if an exception occurs then the error is passed back by reference through this parameter</param>
        ///     ''' <returns></returns>
        public static bool DownloadSingleFile(string ftpAddress, string ftpUser, string ftpPassword, string downloadTargetFolder, bool deleteAfterDownload)
        {
            bool FileDownloaded = false;
            try
            {
                string sFtpFile = ftpAddress;

                // Dim sTargetFileName = System.IO.Path.GetFileName(sFtpFile)
                // sTargetFileName = sTargetFileName.Replace("/", "\")
                // sTargetFileName = downloadTargetFolder & sTargetFileName
                using (WebClient wc = new WebClient())
                {
                    wc.Proxy = null;
                    wc.BaseAddress = "ftp://ftp.balaji.com/incoming/";

                    // Authenticate, then download a file to the FTP server.
                    // The same approach also works for HTTP and HTTPS.

                    wc.Credentials = new NetworkCredential(ftpUser, ftpPassword);

                    wc.DownloadFile("log.txt", downloadTargetFolder);
                }

                //Network.DownloadFile(sFtpFile, downloadTargetFolder, ftpUser, ftpPassword);

                if (deleteAfterDownload)
                {
                    System.Net.FtpWebRequest ftpRequest = null;

                    ftpRequest = (System.Net.FtpWebRequest)System.Net.WebRequest.Create(sFtpFile);

                    {
                        var withBlock = ftpRequest;
                        withBlock.Credentials = new System.Net.NetworkCredential(ftpUser, ftpPassword);
                        withBlock.Method = System.Net.WebRequestMethods.Ftp.DeleteFile;
                    }

                    System.Net.FtpWebResponse response = (System.Net.FtpWebResponse)ftpRequest.GetResponse();
                    response.Close();

                    ftpRequest = null;

                    FileDownloaded = true;
                }
            }
            catch (Exception ex)
            {
            }

            return FileDownloaded;
        }
    }
}
