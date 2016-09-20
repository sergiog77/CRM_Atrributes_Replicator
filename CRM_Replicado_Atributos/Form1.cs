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
using System.IO;



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

          //  var entities = metaDataResponse.EntityMetadata.Where(EntityMetadata => EntityMetadata.IsCustomizable.Value == true).OrderBy(EntityMetadata => EntityMetadata.LogicalName);
            var entities = metaDataResponse.EntityMetadata.Where(EntityMetadata => EntityMetadata.LogicalName == "lms_remesa").OrderBy(EntityMetadata => EntityMetadata.LogicalName);

            


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
                EntityFilters = EntityFilters.Attributes| EntityFilters.Relationships,
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


            StreamWriter file2 = new StreamWriter(@"log.txt", true);
            file2.WriteLine("Actualizando las entidades:");
            file2.Close();




            foreach (TreeNode tn in treeView1.Nodes)
            {
                try
                {
                    if (tn.Checked)
                    {
                        var entities = metaDataResponse.EntityMetadata.Where(EntityMetadata => EntityMetadata.LogicalName == tn.Text).OrderBy(EntityMetadata => EntityMetadata.LogicalName);

                        foreach (EntityMetadata entity in entities)
                        {
                            StreamWriter file1 = new StreamWriter(@"log.txt", true);
                            file1.WriteLine("   Actualizando la entidad " + entity.LogicalName);
                            file1.Close();
                            

                            //get the  atribute name from entity source
                            RetrieveAttributeRequest attributeRequest = new RetrieveAttributeRequest
                            {
                                EntityLogicalName = entity.LogicalName,
                                LogicalName = entity.PrimaryNameAttribute,
                                RetrieveAsIfPublished = true,
                            };
                            RetrieveAttributeResponse attributeResponse =
                         (RetrieveAttributeResponse)_serviceProxy.Execute(attributeRequest);


                            try
                            {

                                CreateEntityRequest createrequest = new CreateEntityRequest
                                {
                                   // HasNotes = entity,
                                   // HasActivities = false,

                                    //Define the entity
                                    Entity = new EntityMetadata
                                    {
                                        SchemaName = entity.SchemaName,
                                        DisplayName = entity.DisplayName,
                                        DisplayCollectionName = entity.DisplayCollectionName,
                                        Description = entity.Description,
                                        OwnershipType = entity.OwnershipType,
                                        IsActivity = entity.IsActivity,
                                        IsAvailableOffline=true

                                    },
                                    //////////// Define the primary attribute for the entity
                                    PrimaryAttribute = new StringAttributeMetadata
                                    {
                                        SchemaName = attributeResponse.AttributeMetadata.SchemaName,
                                        RequiredLevel = attributeResponse.AttributeMetadata.RequiredLevel,
                                        MaxLength = 100,
                                        FormatName = StringFormatName.Text,
                                        DisplayName = attributeResponse.AttributeMetadata.DisplayName,
                                        Description = attributeResponse.AttributeMetadata.Description,
                                    }

                                };
                                _serviceProxyProdware.Execute(createrequest);
                            }
                            catch (Exception ex)
                            {
                                StreamWriter file3 = new StreamWriter(@"log.txt", true);
                                file3.WriteLine("      Error al crear la entidad " + tn.Text + ". Error:" + ex.Message);
                                file3.Close();
                            }
                            CrearAtributossEntidad(entity, tn);

                        }

                    }
                }
                catch (Exception ex)
                {
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (TreeNode tn in treeView1.Nodes)
            {
                tn.Checked = true;

                foreach (TreeNode child in tn.Nodes)
                {
                    child.Checked = true;
                }
            }


        }




        private void CrearAtributossEntidad(EntityMetadata entidad, TreeNode nombreentidad)
        {


            RetrieveEntityRequest entityRequest1 = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.Attributes|EntityFilters.Relationships,
                LogicalName = entidad.LogicalName,
                RetrieveAsIfPublished = true
            };

            RetrieveEntityResponse entityResponse1 = (RetrieveEntityResponse)_serviceProxy.Execute(entityRequest1);
            List<AttributeMetadata> addedAttributes = new List<AttributeMetadata>();
            foreach (TreeNode child in nombreentidad.Nodes)
            {
                if (child.Checked)
                {
                    foreach (AttributeMetadata atributo in entityResponse1.EntityMetadata.Attributes)
                    {
                        try
                        {
                            if (atributo.LogicalName.ToLower() == child.Text.ToLower())
                                addedAttributes.Add(atributo);
                        }
                        catch (Exception ex)
                        {
                        }

                    }
                }
            }


            foreach (AttributeMetadata anAttribute in addedAttributes)
            {
                StreamWriter file1 = new StreamWriter(@"log.txt", true);
                file1.WriteLine("      Actualizando el atributo : " + anAttribute.LogicalName);
                file1.Close();



                try
                {

                    if (anAttribute.GetType().ToString() == "Microsoft.Xrm.Sdk.Metadata.LookupAttributeMetadata")
                    {
                        CrearAtributoLookup(anAttribute, entityResponse1);
                        continue;
                    }

                    else if (anAttribute.GetType().ToString() == "Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata")
                    {
                        if (((Microsoft.Xrm.Sdk.Metadata.OptionSetMetadataBase)(((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)(anAttribute)).OptionSet)).IsGlobal == true)
                        {
                            ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)(anAttribute)).OptionSet.Options.Clear();

                        }
                    }
                    else if (anAttribute.GetType().ToString() == "Microsoft.Xrm.Sdk.Metadata.MemoAttributeMetadata")
                    {

                        MemoAttributeMetadata memoAttribute = new MemoAttributeMetadata
                        {
                            // Set base properties
                            SchemaName = anAttribute.SchemaName,
                            DisplayName = anAttribute.DisplayName,
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                            Description = anAttribute.Description,
                            // Set extended properties
                            Format = Microsoft.Xrm.Sdk.Metadata.StringFormat.TextArea,
                            ImeMode = Microsoft.Xrm.Sdk.Metadata.ImeMode.Disabled,
                            MaxLength = 500
                        };

                        CreateAttributeRequest createAttributeRequest1 = new CreateAttributeRequest
                        {
                            EntityName = entidad.LogicalName,
                            Attribute = memoAttribute
                        };
                        // Execute the request.
                        _serviceProxyProdware.Execute(createAttributeRequest1);






                    }





                    // Create the request.
                    CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                    {
                        EntityName = entidad.LogicalName,
                        Attribute = anAttribute
                    };
                    // Execute the request.
                    _serviceProxyProdware.Execute(createAttributeRequest);

                    StreamWriter file3 = new StreamWriter(@"log.txt", true);
                    file3.WriteLine("      atributo " + anAttribute.SchemaName + ". Creado correctamente");
                    file3.Close();

                }
                catch (Exception ex)
                {
                    StreamWriter file3 = new StreamWriter(@"log.txt", true);
                    file3.WriteLine("      Error al crear el atributo " + anAttribute.SchemaName + ". Error:" + ex.Message);
                    file3.Close();
                }

            }

        }


        private void CrearAtributoLookup(AttributeMetadata anAttribute,  RetrieveEntityResponse entidad )
        {
            try
            {

                var oneToNRelationships = entidad.EntityMetadata.ManyToOneRelationships;


                AssociatedMenuConfiguration menuconf = new AssociatedMenuConfiguration();
                CascadeConfiguration cascade = new CascadeConfiguration();


                foreach (OneToManyRelationshipMetadata relation in oneToNRelationships)
                {
                    if (relation.ReferencingAttribute == anAttribute.SchemaName)
                    {
                        menuconf = relation.AssociatedMenuConfiguration;
                        cascade = relation.CascadeConfiguration;

                        CreateOneToManyRequest createOneToManyRelationshipRequest =
               new CreateOneToManyRequest
               {
                   OneToManyRelationship =
                   new OneToManyRelationshipMetadata
                   {
                       ReferencedEntity = relation.ReferencedEntity,
                       ReferencingEntity = anAttribute.EntityLogicalName,
                       SchemaName = ((Microsoft.Xrm.Sdk.Metadata.RelationshipMetadataBase)(relation)).SchemaName,
                       AssociatedMenuConfiguration = new AssociatedMenuConfiguration
                       {
                           Behavior = menuconf.Behavior,
                           Group = menuconf.Group,
                           Label = menuconf.Label,
                           Order = menuconf.Order,
                       },
                       CascadeConfiguration = new CascadeConfiguration
                       {
                           Assign = cascade.Assign,
                           Delete = cascade.Delete,
                           Merge = cascade.Merge,
                           Reparent = cascade.Reparent,
                           Share = cascade.Share,
                           Unshare = cascade.Unshare
                       }
                   },
                   Lookup = new LookupAttributeMetadata
                   {
                       SchemaName = anAttribute.SchemaName,
                       DisplayName = anAttribute.DisplayName,
                       RequiredLevel = anAttribute.RequiredLevel,
                       Description = anAttribute.Description
                   }
               };
                        var createOneToManyRelationshipResponse = (CreateOneToManyResponse)_serviceProxyProdware.Execute(createOneToManyRelationshipRequest);



                        break;
                    }
                }

            }
            catch (Exception ex)
            { 
            }

           



        }

    }
}
