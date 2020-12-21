using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.TestManagement.Client;

namespace AIT.TFS.SyncService.Contracts.TestCenter
{
    public class TfsSharedStep : ITfsSharedStep
    {
        public TfsSharedStep(ISharedStep originalSharedStep)
        {
            if(originalSharedStep == null)
                throw  new ArgumentNullException("originalSharedStep");
            OrginialSharedStep = originalSharedStep;
            Id = OrginialSharedStep.Id;
            Title = OrginialSharedStep.WorkItem.Title;
        }

        public ISharedStep OrginialSharedStep {get; private set;}

        public int Id { get; private set;}

        public string Title { get; private set;}
    }
}
