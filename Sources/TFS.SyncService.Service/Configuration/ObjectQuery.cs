using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace AIT.TFS.SyncService.Contracts.Configuration
{

    /// <summary>
    /// Implementation of the object query
    /// </summary>
   public  class ObjectQuery : IObjectQuery
    {

       /// <summary>
       /// Constructor
       /// </summary>
       public ObjectQuery()
       {
           DestinationElements = new List<IElement>();
           //WorkItemLinkFilters = new List<IWorkItemLinkFilter>();
       }

       /// <summary>
       /// The name of the query
       /// </summary>
        public string Name
        {
            get;
            set;
        }

       /// <summary>
       /// The target elements which should be contained in the result list
       /// </summary>
        public IList<IElement> DestinationElements
        {
            get;
            set;
        }

       /// <summary>
       ///  The filters 
       /// </summary>

       public IWorkItemLinkFilter WorkItemLinkFilters
       {
           get;
           set;
       }

       /// <summary>
       /// Options for the filter
       /// </summary>
       public IFilterOption FilterOption
       {
           get;
           set;
       }

        /// <summary>
        /// Returns a string that represents the current object.
        /// </summary>
        /// <returns>
        /// A string that represents the current object.
        /// </returns>
        public override string ToString()
        {
            String returnString = "";
            returnString+= $"Name: {Name}:";

            if (DestinationElements != null)
            {
                foreach (var VARIABLE in DestinationElements)
                {
                     returnString+= $"DestinationElement: {VARIABLE.ToString()}:";
                }
              
            }

            if (FilterOption != null)
            {
                returnString += $"FilterOptions:Distinct {FilterOption.Distinct}, Latest {FilterOption.Latest}, Property {FilterOption.FilterProperty}:";
            }

            return returnString;

        }
    }
}
