using System.Collections.Generic;
using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts.Configuration;
using AIT.TFS.SyncService.Contracts.Enums;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{
    public class WorkItemLinkFilter
    {


        [XmlAttribute("FilterType")]
        public FilterType FilterType
        {
            get;
            set;
        }


       
        [XmlElement("Filter")]

        public List<Filter> Filters
        {
            get;
            set;
        }

        public class Filter : IFilter
        {
             [XmlAttribute("LinkType")]
            public string LinkType
            {
                get;
                set;
            }
             [XmlAttribute("FilterOn")]
            public string FilterOn
            {
                get;
                set;
            }
        }

    }
}
