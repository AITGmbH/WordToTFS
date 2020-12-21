using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIT.TFS.SyncService.Contracts.WorkItemObjects
{
    public class WorkItemTypeAppearance
    {

        public WorkItemTypeAppearance(string workItemTypeName, string primaryColor)
        {
            WorkItemTypeName = workItemTypeName;
            PrimaryColor = primaryColor;
        }

        public string PrimaryColor { get; set; }

        public string WorkItemTypeName { get; set; }
    }
}
