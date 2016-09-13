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
        public static Uri oUri = new Uri("https://legdevcr.legalitas.es/lmsmultasfusdes/XRMServices/2011/Organization.svc");
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
            crm.Conexion = new Uri("https://legdevcr.legalitas.es/lmsmultasfusdes/XRMServices/2011/Organization.svc");
            crm.credenciales = clientCredentials;
            _serviceProxy = crm.CrmService();

            RetrieveAllEntitiesRequest metaDataRequest = new RetrieveAllEntitiesRequest();
            RetrieveAllEntitiesResponse metaDataResponse = new RetrieveAllEntitiesResponse();
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

                    foreach (TreeNode child in tn.Nodes)
                    {
                        if (child.Checked)
                        { 
                            string  nombre = child.Text;
                        }

                    }
                }
            }
        }









    }
}
