using Microsoft.SharePoint.Client;
using SPMeta2.Common;
using SPMeta2.CSOM.Services;
using SPMeta2.Definitions;
using SPMeta2.Extensions;
using SPMeta2.Models;
using SPMeta2.Services;
using SPMeta2.Syntax.Default;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace SPF.M2
{
    public class SpMetaEasyDeploy
    {
        private ClientContext _ctx;


        public bool Incremental { get; set; }
        public Action<string> Logger { get; set; }


        public SpMetaEasyDeploy(ClientContext ctx)
        {
            _ctx = ctx;
            Incremental = false;
            Logger = Console.WriteLine;
        }

        public void DeploySPMetaModel(WebModelNode model)
        {
            BeforeDeployModel(x =>
            {
                PropertyBagValue incrementalProvisionModelIdProperty = GetIncrementalProvisionModelProperies(model.PropertyBag);

                Logger("Provisioning preparing model");

                WebModelNode preparingModel = GetContainersModel(model); 

                if (incrementalProvisionModelIdProperty != null)
                {
                    preparingModel.SetIncrementalProvisionModelId("Preparing: " + incrementalProvisionModelIdProperty.Value);
                }

                x.DeployModel(SPMeta2.CSOM.ModelHosts.WebModelHost.FromClientContext(_ctx), preparingModel);

                Logger("");
                Logger("Provisioning main model");

                x.DeployModel(SPMeta2.CSOM.ModelHosts.WebModelHost.FromClientContext(_ctx), model);
            });
        }

        public void DeploySPMetaModel(SiteModelNode model)
        {
            BeforeDeployModel(x =>
            {
                PropertyBagValue incrementalProvisionModelIdProperty = GetIncrementalProvisionModelProperies(model.PropertyBag);

                Logger("Provisioning preparing model");

                SiteModelNode preparingModel = GetContainersModel(model);

                if (incrementalProvisionModelIdProperty != null)
                {
                    preparingModel.SetIncrementalProvisionModelId("Preparing: " + incrementalProvisionModelIdProperty.Value);
                }

                x.DeployModel(SPMeta2.CSOM.ModelHosts.SiteModelHost.FromClientContext(_ctx), preparingModel);

                Logger("");
                Logger("Provisioning main model");

                x.DeployModel(SPMeta2.CSOM.ModelHosts.SiteModelHost.FromClientContext(_ctx), model);
            });

        }

        private PropertyBagValue GetIncrementalProvisionModelProperies(List<PropertyBagValue> propertyBagList)
        {
            PropertyBagValue incrementalProvisionModelIdProperty = propertyBagList.FirstOrDefault(currentPropertyValue =>
                       currentPropertyValue.Name == "_sys.IncrementalProvision.PersistenceStorageModelId");
            if (Incremental && incrementalProvisionModelIdProperty == null)
            {
                new SystemException("Please set incremental provision model id");
            }
            return incrementalProvisionModelIdProperty;
        }

        private void BeforeDeployModel(Action<CSOMProvisionService> Deploy)
        {
            DateTime startedDate = DateTime.Now;

            CSOMProvisionService provisionService = new CSOMProvisionService();

            if (Incremental)
            {
                IncrementalProvisionConfig incProvisionConfig = new IncrementalProvisionConfig();
                incProvisionConfig.AutoDetectSharePointPersistenceStorage = true;
                provisionService.SetIncrementalProvisionMode(incProvisionConfig);
            }

            provisionService.OnModelNodeProcessed += (sender, args) =>
            {
                ModelNodeProcessed(sender, args);
            };

            Deploy(provisionService);

            provisionService.SetDefaultProvisionMode();

            DateTime finishedDate = DateTime.Now;
            TimeSpan executionTime = (finishedDate - startedDate);

            if (executionTime.Days > 0)
            {
                Logger(String.Format("It took us {3} days and {0}:{1}:{2} hours", executionTime.Hours, executionTime.Minutes, executionTime.Seconds, executionTime.Days));
            }
            else
            {
                Logger(String.Format("It took us {0}:{1}:{2} hours", executionTime.Hours, executionTime.Minutes, executionTime.Seconds));
            }

            Logger("");
            Logger("");
        }

        private WebModelNode GetContainersModel(WebModelNode model)
        {
            WebModelNode containersModel = SPMeta2Model.NewWebModel();

            foreach (ModelNode modelNode in model.ChildModels)
            {
                if (modelNode.Value.GetType() == typeof(WebDefinition))
                {
                    containersModel.AddWeb((WebDefinition)modelNode.Value, currentWeb => {
                        GetWebContainersModel(currentWeb, modelNode.ChildModels);
                    });
                }

                if (modelNode.Value.GetType() == typeof(ListDefinition))
                {
                    containersModel.AddList((ListDefinition)modelNode.Value);
                }
            }

            return containersModel;
        }

        private SiteModelNode GetContainersModel(SiteModelNode model)
        {
            SiteModelNode containersModel = SPMeta2Model.NewSiteModel();

            foreach (ModelNode modelNode in model.ChildModels)
            {
                if (modelNode.Value.GetType() == typeof(WebDefinition))
                {
                    containersModel.AddWeb((WebDefinition)modelNode.Value, currentWeb => {
                        GetWebContainersModel(currentWeb, modelNode.ChildModels);
                    });
                }
            }

            return containersModel;
        }

        private WebModelNode GetWebContainersModel(WebModelNode model, Collection<ModelNode> childModels)
        {
            foreach (ModelNode modelNode in childModels)
            {
                if (modelNode.Value.GetType() == typeof(WebDefinition))
                {
                    model.AddWeb((WebDefinition)modelNode.Value, currentWeb => {
                        GetWebContainersModel(currentWeb, modelNode.ChildModels);
                    });
                }

                if (modelNode.Value.GetType() == typeof(ListDefinition))
                {
                    model.AddList((ListDefinition)modelNode.Value);
                }
            }
            return model;
        }

        #region Metodi rifattorizzati

        private void ModelNodeProcessed(object sender, ModelProcessingEventArgs args)
        {
            string nodeDone = AddZeros(args.ProcessedModelNodeCount, 4);
            string nodeRemaining = AddZeros(args.TotalModelNodeCount, 4);
            string percent = AddSpacesBefore(Math.Round(100d * (double)args.ProcessedModelNodeCount / (double)args.TotalModelNodeCount), 3);
            string nodeName = args.CurrentNode.Value.ToString();
            string modelId = args.Model.GetPropertyBagValue(DefaultModelNodePropertyBagValue.Sys.IncrementalProvision.PersistenceStorageModelId);

            string deployMode = "[+]";
            if (Incremental && !args.CurrentNode.GetIncrementalRequireSelfProcessingValue())
            {
                deployMode = "[-]";
            }

            Logger(string.Format("{0}[{1}] [{2}/{3}] - [{4}%] - [{5}] [{6}]", new object[] {
                deployMode,
                modelId,
                nodeDone,
                nodeRemaining,
                percent,
                args.CurrentNode.Value.GetType().Name,
                nodeName
            }));
        }

        private string AddZeros(double number, int zeros)
        {
            return AddBeforeSymbols(number.ToString(), zeros, '0');
        }

        private string AddSpacesBefore(double number, int zeros)
        {
            return AddBeforeSymbols(number.ToString(), zeros, ' ');
        }

        private string AddBeforeSymbols(string strValue, int zeros, char symbol)
        {
            for (var i = 0; i < zeros; i++)
            {
                strValue = symbol + strValue;
            }
            strValue = strValue.Substring(strValue.Length - zeros);

            return strValue;
        }


        #endregion
    }
}
