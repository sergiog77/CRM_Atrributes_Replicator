using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Metadata;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Client;
using CRM_Replicado_Atributos.AppCode;



namespace CRM_Replicado_Atributos
{
    public partial class Form1 : Form
    {
        public static OrganizationServiceProxy _serviceProxy;
        private static OrganizationService _orgService;

        public static OrganizationServiceProxy _serviceProxyProdware;
        private static OrganizationService _orgServiceProdware;
        public static Uri oUri = new Uri("https://legdevcr.legalitas.es/lmsmultasfusdes/XRMServices/2011/Organization.svc");
        public static Uri oUriProdware = new Uri("https://legdevcr.legalitas.es/lmsfusiondes1/XRMServices/2011/Organization.svc");

        public static RetrieveAllEntitiesResponse metaDataResponse;
        ClientCredentials clientCredentials = new ClientCredentials();


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            clientCredentials.UserName.UserName = "sergio.gomez@legalitas.es";
            clientCredentials.UserName.Password = "legalitas2";

            CrmHelper crm = new CrmHelper();
            crm.Conexion = oUri;
            crm.credenciales = clientCredentials;
            _serviceProxy = crm.CrmService();


            crm.Conexion = oUriProdware;
            _serviceProxyProdware = crm.CrmService();



            RetrieveAllEntitiesRequest metaDataRequest = new RetrieveAllEntitiesRequest();
            //RetrieveAllEntitiesResponse metaDataResponse = new RetrieveAllEntitiesResponse();
            metaDataRequest.EntityFilters = EntityFilters.Entity;
            metaDataResponse = (RetrieveAllEntitiesResponse)_serviceProxy.Execute(metaDataRequest);
            metaDataRequest.EntityFilters = EntityFilters.Entity;
            metaDataResponse = (RetrieveAllEntitiesResponse)_serviceProxy.Execute(metaDataRequest);
            //obtenemos el listado de las entiedades de multas

            var entities = metaDataResponse.EntityMetadata.Where(EntityMetadata => EntityMetadata.IsCustomizable.Value == true).OrderBy(EntityMetadata => EntityMetadata.LogicalName);


            foreach (EntityMetadata entity in entities)
            {
                TreeNode nodo = new TreeNode();
                nodo.Text = entity.LogicalName;
                ObtenerAtributosEntidad(entity, nodo);
                treeView1.Nodes.Add(nodo);
                treeView1.EndUpdate();
            }











        }


        public static void ObtenerAtributosEntidad(EntityMetadata entity, TreeNode nodo)
        {

            RetrieveEntityRequest entityRequest1 = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.Attributes,
                LogicalName = entity.LogicalName,
                RetrieveAsIfPublished = true
            };

            RetrieveEntityResponse entityResponse1 = (RetrieveEntityResponse)_serviceProxy.Execute(entityRequest1);


            foreach (AttributeMetadata atributo in entityResponse1.EntityMetadata.Attributes)
            {
                TreeNode nodochild = new TreeNode();
                nodochild.Text = atributo.LogicalName;
                nodo.Nodes.Add(nodochild);

            }




        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (TreeNode tn in treeView1.Nodes)
            {
                if (tn.Checked)
                {
                    var entities = metaDataResponse.EntityMetadata.Where(EntityMetadata => EntityMetadata.LogicalName == tn.Text).OrderBy(EntityMetadata => EntityMetadata.LogicalName);

                    foreach (EntityMetadata entity in entities)
                    {

                        //get the  atribute name from entity source
                        RetrieveAttributeRequest attributeRequest = new RetrieveAttributeRequest
                        {
                            EntityLogicalName = entity.LogicalName,
                            LogicalName =  entity.PrimaryNameAttribute,
                            RetrieveAsIfPublished = true
                        };
                        RetrieveAttributeResponse attributeResponse =
                     (RetrieveAttributeResponse)_serviceProxy.Execute(attributeRequest);


                        CreateEntityRequest createrequest = new CreateEntityRequest
                        {

                            //Define the entity
                            Entity = new EntityMetadata
                            {
                                SchemaName = entity.SchemaName,
                                DisplayName = entity.DisplayName,
                                DisplayCollectionName =  entity.DisplayCollectionName,
                                Description = entity.Description, 
                                OwnershipType = OwnershipTypes.UserOwned,
                                IsActivity = false,

                            },
                            //////////// Define the primary attribute for the entity
                            PrimaryAttribute = new StringAttributeMetadata
                            {
                                SchemaName = attributeResponse.AttributeMetadata.SchemaName,
                                RequiredLevel = attributeResponse.AttributeMetadata.RequiredLevel,
                                MaxLength = 100,
                                FormatName = StringFormatName.Text,
                                DisplayName =  attributeResponse.AttributeMetadata.DisplayName,
                                Description = attributeResponse.AttributeMetadata.Description,
                            }

                        };
                        _serviceProxyProdware.Execute(createrequest);





                        RetrieveEntityRequest entityRequest1 = new RetrieveEntityRequest
                        {
                            EntityFilters = EntityFilters.Attributes,
                            LogicalName = entity.LogicalName,
                            RetrieveAsIfPublished = true
                        };

                        RetrieveEntityResponse entityResponse1 = (RetrieveEntityResponse)_serviceProxy.Execute(entityRequest1);

                        List<AttributeMetadata>   addedAttributes = new List<AttributeMetadata>();
                        foreach (AttributeMetadata atributo in entityResponse1.EntityMetadata.Attributes)
                        {
                            addedAttributes.Add(atributo);

                        }




                        foreach (AttributeMetadata anAttribute in addedAttributes)
                        {

                            try
                            {
                                // Create the request.
                                CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                                {
                                    EntityName = entity.LogicalName,
                                    Attribute = anAttribute
                                };

                                // Execute the request.

                                _serviceProxyProdware.Execute(createAttributeRequest);
                            }
                            catch (Exception ex)
                            { 
                            }

                        }







                    }







                 








                }
            }

        }

    }
}
