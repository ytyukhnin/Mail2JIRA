using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace Mail2JIRA.JiraDTO
{
    [JsonObject(MemberSerialization.OptOut)]
    public class CreateIssueResponseObject
    {
        public string Id { get; set; }
        public string Key { get; set; }
        public string Self { get; set; }
    }
}
