using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace Mail2JIRA.JiraDTO
{
    [JsonObject(MemberSerialization.OptOut)]
    public class FieldsObject
    {
        public ProjectObject Project { get; set; }

        public string Summary { get; set; }

        public string Description { get; set; }

        public IssueTypeObject IssueType { get; set; }

        public UserObject Assignee { get; set; }
    }
}
