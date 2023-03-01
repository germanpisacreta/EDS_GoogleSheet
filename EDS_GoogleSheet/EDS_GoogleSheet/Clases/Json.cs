using System;
using System.Collections.Generic;
using System.Text;

namespace EDS_GoogleSheet.Clases
{
    public class Json
    {
        public string module { get; set; }
        public string sequenceid { get; set; }
        public string sessionid { get; set; }
        public string conversationid { get; set; }
        public string evgid { get; set; }
        public string callid { get; set; }
        public string startdatetime { get; set; }
        public string enddatetime { get; set; }
        public string duration { get; set; }
        public string channel { get; set; }
        public string user { get; set; }
        public string username { get; set; }
        public string access { get; set; }
        public string rn { get; set; }
        public string releaseside { get; set; }
        public string resultcode { get; set; }
        public string segment { get; set; }
        public string firstuc { get; set; }
        public string lastuc { get; set; }
        public string openpromptqtty { get; set; }
        public string highscorerspqtty { get; set; }
        public string lowscorerspqtty { get; set; }
        public string confyesqtty { get; set; }
        public string confnoqtty { get; set; }
        public string confotherqtty { get; set; }
        public string longrespqtty { get; set; }
        public string nuisanceqtty { get; set; }
        public string othersqtty { get; set; }
        public string silencenoiseqtty { get; set; }
        public string firstintentid { get; set; }
        public string firstintentscore { get; set; }
        public string lastintentid { get; set; }
        public string lastintentscore { get; set; }
        public string trace { get; set; }
        public string survey { get; set; }
        public string surveyFreeText { get; set; }
        public string id { get; set; }
        public string _rid { get; set; }
        public string _self { get; set; }
        public string _etag { get; set; }
        public string _attachments { get; set; }
        public string _ts { get; set; }
        public List<questionActionRecords> questionActionRecords { get; set; }

    }
}
