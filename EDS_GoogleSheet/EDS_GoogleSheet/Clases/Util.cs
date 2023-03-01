using Google.Apis.Drive.v3;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json;

namespace EDS_GoogleSheet.Clases
{
    public static class Util
    {
        static StreamWriter sw = null;
        public static void ListarArchivos(DriveService service, List<String> lstFiles, List<Json> lstJson)
        {
            try
            {
                string path;
                //LISTAR ARCHIVOS DEL DRIVE
                //  Definir lo parámetros de solicitud
                FilesResource.ListRequest listRequest = service.Files.List();
                listRequest.PageSize = 500;
                listRequest.Fields = "nextPageToken, files(id, name)";

                // Mostramos la lista de archivos
                IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                    .Files;
                if (files != null && files.Count > 0)
                {
                    foreach (var archivos in files.Where(x => x.Name.Contains(".json")).ToList())
                    {
                        var getRequest = service.Files.Get(archivos.Id);
                        var FileStream = new FileStream(archivos.Name, FileMode.Create, FileAccess.Write);
                        getRequest.Download(FileStream);
                        lstFiles.Add(archivos.Name);
                        //descargarArchivo(archivos.Id, service);
                    }
                }

                path = MakePath();
                DirectoryInfo di = new DirectoryInfo(path);
                foreach (var fi in di.GetFiles())
                {
                    if (fi.Name.Contains(".json"))
                    {
                        if (lstFiles.Contains(fi.Name))
                        {
                            if (fi.Length > 0)
                            {


                                using (StreamReader jsonStream = System.IO.File.OpenText(fi.ToString()))
                                {
                                    var json = jsonStream.ReadToEnd();
                                    Json product = JsonConvert.DeserializeObject<Json>(json);
                                    lstJson.Add(product);
                                }
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ListarArchivos - " + ex.Message);
            }
        }


        //DESCARGAR UN ARCHIVO DEL DRIVE
        private static void descargarArchivo(string idArchivo, DriveService service)
        {
            try
            {
                var requestSheet = service.Files.List();
                var responseSheet = requestSheet.Execute();

                if (responseSheet.Files.Count > 0)
                {
                    var downloadFile = responseSheet.Files.FirstOrDefault(files => files.Id == idArchivo);
                    var getRequest = service.Files.Get(downloadFile.Id);
                    var FileStream = new FileStream(downloadFile.Name, FileMode.Create, FileAccess.Write);
                    getRequest.Download(FileStream);


                }

            }
            catch (Exception ex)
            {
                throw new Exception("descargarArchivo - " + ex.Message);
            }

        }

        public static string MakePath()
        {
            try
            {
                string pathAux = Assembly.GetExecutingAssembly().Location;
                String[] array;
                array = pathAux.Split("\\");
                string path = null;
                for (int i = 0; i < array.Length - 1; i++)
                {
                    path += array[i] + "\\";
                }
                //le quitamos la ultima ,(coma) a la condicion
                path = path.Remove(path.LastIndexOf('\\'));
                return path;
            }
            catch (Exception ex)
            {
                throw new Exception("ListarArchivos - " + ex.Message);
            }


        }

        public static void DeleteFile(List<String> lstFiles)
        {
            try
            {
                string path = MakePath();
                DirectoryInfo di = new DirectoryInfo(path);
                foreach (var fi in di.GetFiles())
                {
                    if (lstFiles.Contains(fi.Name))
                    {
                        System.IO.File.Delete(fi.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            
         
        }

        #region LogFile
        public  static void OpenLogFile(string Name, string path)
        {
            try
            {
                //Si no existe la carpeta Log en el mismo lugar donde se ejecuta el fuente la creo
                if (Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                string pathFinal = string.Format(@"{0}\{1}", path, Name + DateTime.Now.ToString("ddMMyyyyHHmm") + ".txt");

                if (File.Exists(pathFinal))
                    File.Delete(pathFinal);
                sw = new StreamWriter(pathFinal);
            }
            catch (Exception ex)
            {
                throw new Exception("OpenLogFile - " + ex.Message);
            }
        }

       

        public static void WriteLogFileErrorLine(string texto)
        {
            try
            {
                sw.WriteLine(texto);
            }
            catch (Exception ex)
            {
                throw new Exception("WriteLogFile - " + ex.Message);
            }
        }

        public static void CloseLogFile()
        {
            try
            {
                sw.Close();
                sw.Dispose();
            }
            catch (Exception ex)
            {
                throw new Exception("CloseLogFile - " + ex.Message);
            }
        }
        #endregion


    }
}
