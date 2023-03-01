using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Net.Security;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EDS_GoogleSheet.Clases
{
    public class SheetsClient
    {
        #region Propiedades
        static List<IList<Object>> objNewRecords = new List<IList<Object>>();
        #endregion

        #region Contructor
        public SheetsClient()
        {

        }
        #endregion

        #region Metodos
        public void UpdateSpreadSheet(List<Json> lstJson)
        {
            try
            {
                var SheetId = "12H05py3oLZiVKbA0eCxEhF2iK2q-zq-j8dGRbpfBLtQ";
                var service = AuthorizeGoogleAppForSheetsService();
                string newRange = GetRange(service, SheetId);
                IList<IList<Object>> objNeRecords = GenerateData(lstJson);
                UpdatGoogleSheetinBatch(objNeRecords, SheetId, newRange, service);
            }
            catch (Exception ex)
            {
                throw new Exception("UpdateSpreadSheet - : " + ex.Message);
            }
            
        }

        private static SheetsService AuthorizeGoogleAppForSheetsService()
        {
            // If modifying these scopes, delete your previously saved credentials  
            // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json  
            string[] Scopes = { SheetsService.Scope.Spreadsheets };
            string ApplicationName = "EDS-GoogleSheet";
            UserCredential credential;
            using (var stream =
               new FileStream("credenciales.json", FileMode.Open, FileAccess.Read))
            {

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore("MyAppsToken")).Result;

            }

            // Create Google Sheets API service.  
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }
        

        protected static string GetRange(SheetsService service, string SheetId)
        {
            try
            {
                // Define request parameters.  
                String spreadsheetId = SheetId;
                //String range = "A:A";
                string sheetName = DateTime.Now.ToString("ddMMyyyyHHmm");
                var addSheetRequest = new AddSheetRequest();
                addSheetRequest.Properties = new SheetProperties();
                addSheetRequest.Properties.Title = sheetName;
                BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
                batchUpdateSpreadsheetRequest.Requests = new List<Request>();
                batchUpdateSpreadsheetRequest.Requests.Add(new Request { AddSheet = addSheetRequest });

                var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId);
                batchUpdateRequest.Execute();

                
                String range =  sheetName ;

                SpreadsheetsResource.ValuesResource.GetRequest getRequest =
                           service.Spreadsheets.Values.Get(spreadsheetId, range);
                System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                ValueRange getResponse = getRequest.Execute();
                IList<IList<Object>> getValues = getResponse.Values;
                if (getValues == null)
                {
                    return $"{sheetName}!A:A";
                    // spreadsheet is empty return Row A Column A  
                    //return "A" + sheetName + ":A";
                }

                int currentCount = getValues.Count() + 1;
                String newRange = "A" + currentCount + ":A";
                return newRange;
            }
            catch (Exception ex)
            {
                throw new Exception("GetRange - : " + ex.Message);
            }
            
        }

        private static IList<IList<Object>> GenerateData(List<Json> lstJson)
        {
            try
            {
                objNewRecords = new List<IList<Object>>();
                int maxrows = lstJson.Count;
                PropertyInfo[] lstPropiedades = typeof(Json).GetProperties();
                PropertyInfo[] lstPropiedadesquestionActionRecords = typeof(questionActionRecords).GetProperties();
                IList<Object> obj = new List<Object>();
                foreach (var item in lstPropiedades)
                {

                    if (item.Name != "questionActionRecords")
                    {
                        obj.Add(item.Name);
                    }
                    else
                    {
                        foreach (var itemquestionActionRecords in lstPropiedadesquestionActionRecords)
                        {
                            obj.Add(itemquestionActionRecords.Name);
                        }
                    }

                }
                objNewRecords.Add(obj);
                for (var i = 0; i < maxrows; i++)
                {
                    if (lstJson[i] != null)
                    {
                        if (lstJson[i].questionActionRecords != null)
                        {
                            if (lstJson[i].questionActionRecords.Count > 0)
                            {
                                foreach (var questionActionRecords in lstJson[i].questionActionRecords)
                                {
                                    CreateCells(lstJson[i], questionActionRecords);
                                }
                            }
                            else
                            {
                                CreateCells(lstJson[i], null);

                            }
                        }
                        else
                        {
                            CreateCells(lstJson[i], null);
                        }

                    }

                }
                return objNewRecords;

            }
            catch (Exception ex)
            {
                throw new Exception("GenerateData - " + ex.Message);
            }
        }

        private static void CreateCells(Json objJson, questionActionRecords questionActionRecords = null)
        {
            try
            {
                IList<Object> obj = new List<Object>();
                obj.Add(objJson.module);
                obj.Add(objJson.sequenceid);
                obj.Add(objJson.sessionid);
                obj.Add(objJson.conversationid);
                obj.Add(objJson.evgid);
                obj.Add(objJson.callid);
                obj.Add(objJson.startdatetime);
                obj.Add(objJson.enddatetime);
                obj.Add(objJson.duration);
                obj.Add(objJson.channel);
                obj.Add(objJson.user);
                obj.Add(objJson.username);
                obj.Add(objJson.access);
                obj.Add(objJson.rn);
                obj.Add(objJson.releaseside);
                obj.Add(objJson.resultcode);
                obj.Add(objJson.segment);
                obj.Add(objJson.firstuc);
                obj.Add(objJson.lastuc);
                obj.Add(objJson.openpromptqtty);
                obj.Add(objJson.highscorerspqtty);
                obj.Add(objJson.lowscorerspqtty);
                obj.Add(objJson.confyesqtty);
                obj.Add(objJson.confnoqtty);
                obj.Add(objJson.confotherqtty);
                obj.Add(objJson.longrespqtty);
                obj.Add(objJson.nuisanceqtty);
                obj.Add(objJson.othersqtty);
                obj.Add(objJson.silencenoiseqtty);
                obj.Add(objJson.firstintentid);
                obj.Add(objJson.firstintentscore);
                obj.Add(objJson.lastintentid);
                obj.Add(objJson.lastintentscore);
                obj.Add(objJson.trace);
                obj.Add(objJson.survey);
                obj.Add(objJson.surveyFreeText);
                obj.Add(objJson.id);
                obj.Add(objJson._rid);
                obj.Add(objJson._self);
                obj.Add(objJson._etag);
                obj.Add(objJson._attachments);
                obj.Add(objJson._ts);
                if (questionActionRecords != null)
                {
                    obj.Add(questionActionRecords.question);
                    obj.Add(questionActionRecords.activationKind);
                    obj.Add(questionActionRecords.luisScore);
                    obj.Add(questionActionRecords.useCaseName);
                    obj.Add(questionActionRecords.qnaInterKbScore);
                    obj.Add(questionActionRecords.qnaIntraKbScore);
                    obj.Add(questionActionRecords.qnaAnswer);
                    obj.Add(questionActionRecords.qnaAnswerId);
                }
                objNewRecords.Add(obj);
            }
            catch (Exception ex)
            {
                throw new Exception("CreateCells - " + ex.Message);
            }
           

        }


        private static void UpdatGoogleSheetinBatch(IList<IList<Object>> values, string spreadsheetId, string newRange, SheetsService service)
        {
            try
            {
                SpreadsheetsResource.ValuesResource.AppendRequest request =
              service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, spreadsheetId, newRange);
                request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                var response = request.Execute();
            }
            catch (Exception ex)
            {
                throw new Exception("UpdatGoogleSheetinBatch - " + ex.Message);
            }
           
        }
        #endregion
    }
}