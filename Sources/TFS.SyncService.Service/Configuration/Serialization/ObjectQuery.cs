using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Xml.Serialization;
using AIT.TFS.SyncService.Contracts.Configuration;

namespace AIT.TFS.SyncService.Service.Configuration.Serialization
{

    /// <summary>
    /// Class that represents a serialzable class of the object query
    /// </summary>
    public class ObjectQuery
    {

        /// <summary>
        /// Property used to serialize 'name' xml attribute.
        /// </summary>
        [XmlAttribute("Name")]
        public string Name { get; set; }

        /// <summary>
        /// The options that can be provided
        /// </summary>
        [XmlElement("Options")]
        public FilterOption FilterOption
        {
            get;
            set;
        }

        /// <summary>
        /// The items that should be contained in the returning list
        /// </summary>
        [XmlArray("DestinationElements")]
        [XmlArrayItem("DestinationElement")]
        public List<Element> DestinationElements
        {
            get;
            set;
        }


        /// <summary>
        /// the list of filters
        /// </summary>
        [XmlElement("WorkItemLinkFilters")]
        public WorkItemLinkFilter WorkItemLinkFilters;

    }

    /// <summary>
    /// the filter options of a object query
    /// </summary>
    public class FilterOption
    {

        /// <summary>
        /// The distinct option will make sure that a distinct list is returned
        /// </summary>
        [XmlAttribute("Distinct")]
        [DefaultValue(false)]
        public Boolean Distinct
        {
            get;
            set;
        }

        /// <summary>
        /// The latest options will filter only for the newest test management object
        /// </summary>
        [XmlAttribute("Latest")]
        [DefaultValue(false)]
        public Boolean Latest
        {
            get;
            set;
        }

        /// <summary>
        /// The FilterProperty names the property which is used to filter the results
        /// </summary>
        [XmlAttribute("FilterProperty")]
        public string FilterProperty
        {
            get;
            set;
        }
        
    }

    public class Element : IElement
    {
        [XmlAttribute("WorkItemType")]
        public string ItemName
        {
            get;
            set;
        }
    }
}
