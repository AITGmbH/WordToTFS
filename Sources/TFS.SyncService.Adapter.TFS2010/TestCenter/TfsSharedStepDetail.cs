using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AIT.TFS.SyncService.Contracts.TestCenter;
using Microsoft.TeamFoundation.TestManagement.Client;
// ReSharper disable All

namespace AIT.TFS.SyncService.Adapter.TFS2012.TestCenter
{
    public class TfsSharedStepDetail : TfsPropertyValueProvider, ITfsSharedStepDetail
    {
        public TfsSharedStepDetail(TfsSharedStep sharedStep)
        {

            if (sharedStep == null)
                throw new ArgumentNullException("sharedStep");

            SharedStepClass = sharedStep;
            Id = SharedStepClass.Id;

            Title = SharedStepClass.OrginialSharedStep.Title;
            if (SharedStepClass.OrginialSharedStep.WorkItem != null)
            {
                IterationPath = SharedStepClass.OrginialSharedStep.WorkItem.IterationPath;
                AreaPath = SharedStepClass.OrginialSharedStep.WorkItem.AreaPath;
                WorkItemId = SharedStepClass.OrginialSharedStep.WorkItem.Id;
            }
        }

        public TfsSharedStep SharedStepClass { get; private set; }

        public ISharedStep OriginalSharedStep
        {
            get { return SharedStepClass.OrginialSharedStep; }
        }

        public IList<TfsTestAction> Actions
        {
            get
            {
                var stepCounter = 1;
                var list = new List<TfsTestAction>();

                //TODO USE CONFIG for following:

                foreach (var action in SharedStepClass.OrginialSharedStep.Actions)
                {
                    list.Add(new TfsTestAction(action, (stepCounter++).ToString(CultureInfo.InvariantCulture), false));
                }
                return list;
            }
        }

        #region Implementation of Interfaces

        public override object AssociatedObject
        {
            get
            {
                return OriginalSharedStep;
            }
        }

        public ITfsSharedStep SharedStep
        {
            get { return SharedStepClass; }
        }

        public int Id { get; private set; }

        public string Title { get; private set; }

        public string IterationPath { get; private set; }

        public string AreaPath { get; private set; }

        public int WorkItemId { get; private set; }

        public ITestCaseParameters TestParametersWithAllValues
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        #endregion Implementation of ITfsSharedStepDetail
    }
}
