using System.Collections.Generic;


namespace AIT.TFS.SyncService.Contracts.Configuration
{

    /// <summary>
    /// A object query that helps to retrieve additional information from the test managment
    /// </summary>
   public  class ObjectQuery : IObjectQuery
    {

       /// <summary>
       /// Public constructor of an object query
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
    }
}
