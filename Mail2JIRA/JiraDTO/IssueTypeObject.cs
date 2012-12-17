using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace Mail2JIRA.JiraDTO
{
    [JsonObject(MemberSerialization.OptOut)]
    public class IssueTypeObject
    {
        public string Name { get; set; }
    }
}
