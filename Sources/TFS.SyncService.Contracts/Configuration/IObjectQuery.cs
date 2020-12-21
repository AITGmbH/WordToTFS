using System;
using System.Collections.Generic;
using AIT.TFS.SyncService.Contracts.Enums;

namespace AIT.TFS.SyncService.Contracts.Configuration
{

    /// <summary>
    /// Interface that describes an object query which is used to get linked worktitems to objects of the testmanagment
    /// </summary>
    public interface IObjectQuery
    {

        /// <summary>
        /// The name of the query
        /// </summary>
        string Name
        {
            get;
        }

        /// <summary>
        /// The destination elements which should be returned by the object query
        /// </summary>
       IList<IElement> DestinationElements     {
            get;
       }


        /// <summary>
        /// The filters that should be applied to the object query
        /// </summary>
        IWorkItemLinkFilter WorkItemLinkFilters
        {
            get;
        }

        /// <summary>
        /// The options that should be applied to the objectquery
        /// </summary>
        IFilterOption FilterOption
        {
            get;
        }

        /// <summary>
        /// ToString method for debug purposes
        /// </summary>
        /// <returns></returns>
        String ToString();


    }


    /// <summary>
    /// The interface that describes the options that are applied to a object query
    /// </summary>
    public interface IFilterOption
    {
        /// <summary>
        /// Return a distinct list
        /// </summary>
        Boolean Distinct
        {
            get;
            set;
        }


        /// <summary>
        /// Return only latest results
        /// </summary>
        Boolean Latest
        {
            get;
            set;
        }

        /// <summary>
        /// Return filtered results
        /// </summary>
        String FilterProperty
        {
            get;
            set;
        }
    }

    /// <summary>
    /// Describes an element in the destination elements
    /// </summary>
    public interface IElement

    {
        /// <summary>
        /// The name of the 
        /// </summary>
        string ItemName
        {
            get;
        }
    }


    /// <summary>
    /// Describes the different 
    /// </summary>
    public interface IWorkItemLinkFilter
    {

        /// <summary>
        /// The type of the filter. Include or Exclude
        /// </summary>
        FilterType FilterType
        {
            get;
        }

        /// <summary>
        /// The various filters to which the filtertype should be applied to 
        /// </summary>
        IList<IFilter> Filters
        {
            get;
        } 
    }

    /// <summary>
    /// The filter used for the link filtering
    /// </summary>
    public interface IFilter
    {

        /// <summary>
        /// The type that should be filtered upon
        /// </summary>
        string LinkType
        {
            get;
        }

    }
}
