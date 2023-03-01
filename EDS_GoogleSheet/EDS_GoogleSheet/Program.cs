using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using Google.Apis.Download;
using Google.Apis.Requests;
using Google;
using Google.Apis.Drive;
using Google.Apis;
using Newtonsoft.Json;
using EDS_GoogleSheet.Clases;
using System.Reflection;

namespace EDS_GoogleSheet
{
    class Program
    {
        static string[] Scopes = { DriveService.Scope.DriveReadonly };
        static string ApplicationName = "EDS-GoogleSheet";
        static List<Json> lstJson = new List<Json>();
        static UserCredential credential;
        static List<String> lstFiles = new List<String>();
       
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Obteniendo credenciales ... ");
                //Defino la Conexion
                DefinirConexion();

                // Crea el servicio de API Drive
                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
                Console.WriteLine("Obteniendo Datos ... ");
                //Listo los archivos del Drive y los guardo en una carpeta para luego leerlos 
                Util.ListarArchivos(service,lstFiles,lstJson);
                Console.WriteLine("Grabando Archivo ... ");
                //Grabo el archivo Excel
                SheetsClient objSheetsClient = new SheetsClient();
                objSheetsClient.UpdateSpreadSheet(lstJson);
                Util.DeleteFile(lstFiles);
                
                Console.WriteLine("Proceso Finalizado\nPresione ENTER para continuar");
                Console.Read();
            }
            catch (Exception ex)
            {
                string path = Util.MakePath();
                Util.OpenLogFile("Error", path);
                Util.WriteLogFileErrorLine(ex.Message);
                Util.CloseLogFile();
                Console.WriteLine("Error ejecutando el proceso revise el log en : " + path + "\nPresione ENTER para continuar");
                Console.Read();
            }
          
        }

        public static void DefinirConexion()
        {
            try
            {
                using (var stream = new FileStream("credenciales.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".token.json");

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                         GoogleClientSecrets.Load(stream).Secrets,
                         Scopes,
                         "user",
                         CancellationToken.None,
                         new FileDataStore(credPath, true)).Result;
                  
                }


               
            }
            catch (Exception ex)
            {
                throw new Exception("DefinirConexion - " + ex.Message);
            }
        }

       
    }
}
