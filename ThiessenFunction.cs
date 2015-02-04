using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using ESRI.ArcGIS.ADF.CATIDs;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.Geoprocessor;
using ESRI.ArcGIS.DataManagementTools;

namespace GPThiessen
{
    public class ThiessenFunction : IGPFunction2
    {
        private string m_ToolName = "Thiessen";
        private IArray m_Parameters; // Array of Parameters
        private IGPUtilities m_GPUtilities; // GPUtilities object        

        public ThiessenFunction()
        {
            m_GPUtilities = new GPUtilitiesClass();
        }

        #region IGPFunction2 Members

        // Set the name of the function tool. 
        // This name appears when executing the tool at the command line or in scripting. 
        // This name should be unique to each toolbox and must not contain spaces.
        public string Name
        {
            get { return m_ToolName; }
        }

        // Set the function tool Display Name as seen in ArcToolbox.
        public string DisplayName
        {
            get { return "Thiessen Polygons"; }
        }

        // This is the location where the parameters to the Function Tool are defined. 
        // This property returns an IArray of parameter objects (IGPParameter). 
        // These objects define the characteristics of the input and output parameters. 
        public IArray ParameterInfo
        {
            get
            {
                // Array to the hold the parameters.
                IArray parameters = new ArrayClass();

                // Input Features
                IGPParameterEdit3 inputParameter = new GPParameterClass();
                inputParameter.DataType = new GPFeatureLayerTypeClass();
                inputParameter.Value = new GPFeatureLayerClass();
                inputParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput;
                inputParameter.DisplayName = "Input Features";
                inputParameter.Name = "input_features";
                inputParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired;
                parameters.Add(inputParameter);

                // Clipping Polygon
                IGPParameterEdit3 polygonParameter = new GPParameterClass();
                polygonParameter.DataType = new GPFeatureLayerTypeClass();
                polygonParameter.Value = new GPFeatureLayerClass();
                polygonParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput;
                polygonParameter.DisplayName = "Clipping Polygon";
                polygonParameter.Name = "clipping_polygon";
                polygonParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired;
                parameters.Add(polygonParameter);

                // Output Features
                IGPParameterEdit3 outputParameter = new GPParameterClass();
                outputParameter.DataType = new DEFeatureClassTypeClass();
                outputParameter.Value = new DEFeatureClassClass();
                outputParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionOutput;
                outputParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired;
                outputParameter.Name = "output_features";
                outputParameter.DisplayName = "Output Feature Class";
                outputParameter.AddDependency("input_features");

                IGPFeatureSchema outSchema = new GPFeatureSchemaClass();
                outSchema.FieldsRule = esriGPSchemaFieldsType.esriGPSchemaFieldsNone;
                outSchema.FeatureTypeRule = esriGPSchemaFeatureType.esriGPSchemaFeatureFirstDependency;
                outSchema.GeometryType = esriGeometryType.esriGeometryPolygon;

                IGPSchema schema = outSchema as IGPSchema;
                schema.GenerateOutputCatalogPath = true;
                outputParameter.Schema = (IGPSchema)outSchema;
                parameters.Add(outputParameter);

                // Unique Field Parameter
                IGPParameterEdit3 fieldParameter = new GPParameterClass();
                fieldParameter.DataType = new FieldTypeClass();
                fieldParameter.Value = new FieldClass();
                fieldParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput;
                fieldParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired;
                fieldParameter.Name = "unique_field";
                fieldParameter.DisplayName = "Unique Field";
                fieldParameter.AddDependency("input_features");
                parameters.Add(fieldParameter);

                return parameters;
            }
        }

        public IGPMessages Validate(IArray paramvalues, bool updateValues, IGPEnvironmentManager envMgr)
        {
            if (m_Parameters == null)
                m_Parameters = ParameterInfo;

            if (updateValues == true)
                UpdateParameters(paramvalues, envMgr);

            IGPMessages validateMsgs = m_GPUtilities.InternalValidate(m_Parameters, paramvalues, updateValues, true, envMgr);

            UpdateMessages(paramvalues, envMgr, validateMsgs);

            return validateMsgs;
        }

        /*
        Called each time the user changes a parameter in the tool dialog or Command Line. 
        This updates the output data of the tool, which extremely useful for building models.  
        After returning from UpdateParameters(), geoprocessing calls its internal validation routine checkng that a given set of parameter values 
        are of the appropriate number, DataType, and value.
        This method will update the output parameter value with the unique field.
        */
        public void UpdateParameters(IArray paramvalues, IGPEnvironmentManager pEnvMgr)
        {
            m_Parameters = paramvalues;

            // Retrieve the input parameter value
            IGPValue parameterValue = m_GPUtilities.UnpackGPValue(m_Parameters.get_Element(0));

            // Retrieve the unique field parameter value
            IGPParameter3 fieldNameParameter = (IGPParameter3)paramvalues.get_Element(3);

            // Get the output feature class schema and empty the additional fields. This will ensure 
            // you don't get dublicate entries.
            IGPParameter3 outputFeatures = (IGPParameter3)paramvalues.get_Element(2);
            IGPFeatureSchema schema = (IGPFeatureSchema)outputFeatures.Schema;
            schema.AdditionalFields = null;

            // If we have an unique field value, create a new field based on the unique field name the user entered.            
            if (fieldNameParameter.Value.IsEmpty() == false)
            {
                string fieldName = fieldNameParameter.Value.GetAsText();

                IField uniqueField = m_GPUtilities.FindField(parameterValue, fieldName);
                IFieldsEdit fieldsEdit = new FieldsClass();
                fieldsEdit.AddField(uniqueField);

                IFields fields = fieldsEdit as IFields;
                schema.AdditionalFields = fields;
            }
        }

        /*
        Called after returning from the update parameters routine. 
        You can examine the messages created from internal validation and change them if desired. 
        */
        public void UpdateMessages(IArray paramvalues, IGPEnvironmentManager pEnvMgr, IGPMessages Messages)
        {
            return;
        }

        public void Execute(IArray paramvalues, ITrackCancel trackcancel, IGPEnvironmentManager envMgr, IGPMessages message)
        {
            // Get Parameters
            IGPParameter3 inputParameter = (IGPParameter3)paramvalues.get_Element(0);
            IGPParameter3 polygonParameter = (IGPParameter3)paramvalues.get_Element(1);
            IGPParameter3 outputParameter = (IGPParameter3)paramvalues.get_Element(2);
            IGPParameter3 fieldParameter = (IGPParameter3)paramvalues.get_Element(3);

            // UnPackGPValue. This ensures you get the value from either the dataelement or GpVariable (ModelBuilder)
            IGPValue inputParameterValue = m_GPUtilities.UnpackGPValue(inputParameter);
            IGPValue polygonParameterValue = m_GPUtilities.UnpackGPValue(polygonParameter);
            IGPValue outputParameterValue = m_GPUtilities.UnpackGPValue(outputParameter);
            IGPValue fieldParameterValue = m_GPUtilities.UnpackGPValue(fieldParameter);

            // Decode Input Feature Layers
            IFeatureClass inputFeatureClass;
            IFeatureClass polygonFeatureClass;

            IQueryFilter inputFeatureClassQF;
            IQueryFilter polygonFeatureClassQF;

            m_GPUtilities.DecodeFeatureLayer(inputParameterValue, out inputFeatureClass, out inputFeatureClassQF);
            m_GPUtilities.DecodeFeatureLayer(polygonParameterValue, out polygonFeatureClass, out polygonFeatureClassQF);

            if (inputFeatureClass == null)
            {
                message.AddError(2, "Could not open input dataset.");
                return;
            }

            if (polygonFeatureClass == null)
            {
                message.AddError(2, "Could not open clipping polygon dataset.");
                return;
            }

            if (polygonFeatureClass.FeatureCount(null) > 1)
            {
                message.AddWarning("Clipping polygon feature class contains more than one feature.");
            }

            // Create the Geoprocessor
            Geoprocessor gp = new Geoprocessor();

            // Create Output Polygon Feature Class
            CreateFeatureclass cfc = new CreateFeatureclass();
            IName name = m_GPUtilities.CreateFeatureClassName(outputParameterValue.GetAsText());
            IDatasetName dsName = name as IDatasetName;
            IFeatureClassName fcName = dsName as IFeatureClassName;
            IFeatureDatasetName fdsName = fcName.FeatureDatasetName as IFeatureDatasetName;

            // Check if output is in a FeatureDataset or not. Set the output path parameter for CreateFeatureClass tool.
            if (fdsName != null)
            {
                cfc.out_path = fdsName;
            }
            else
            {
                cfc.out_path = dsName.WorkspaceName.PathName;
            }

            // Set the output Coordinate System for CreateFeatureClass tool.
            // ISpatialReference3 sr = null;
            IGPEnvironment env = envMgr.FindEnvironment("outputCoordinateSystem");
            
            // Same as Input
            if (env.Value.IsEmpty())
            {
                IGeoDataset ds = inputFeatureClass as IGeoDataset;
                cfc.spatial_reference = ds.SpatialReference as ISpatialReference3;
            }
            // Use the environment setting
            else
            {
                IGPCoordinateSystem cs = env.Value as IGPCoordinateSystem;
                cfc.spatial_reference = cs.SpatialReference as ISpatialReference3;
            }
                       
            // Remaining properties for Create Feature Class Tool
            cfc.out_name = dsName.Name;
            cfc.geometry_type = "POLYGON";

            // Execute Geoprocessor
            gp.Execute(cfc, null);

            // Get Unique Field      
            int iField = inputFeatureClass.FindField(fieldParameterValue.GetAsText());
            IField uniqueField = inputFeatureClass.Fields.get_Field(iField);

            // Extract Clipping Polygon Geometry
            IFeature polygonFeature = polygonFeatureClass.GetFeature(0);
            IPolygon clippingPolygon = (IPolygon)polygonFeature.Shape;

            // Spatial Filter
            ISpatialFilter spatialFilter = new SpatialFilterClass();
            spatialFilter.Geometry = polygonFeature.ShapeCopy;
            spatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelContains;

            // Debug Message
            message.AddMessage("Generating TIN...");

            // Create TIN
            ITinEdit tinEdit = new TinClass();

            // Advanced TIN Functions
            ITinAdvanced2 tinAdv = (ITinAdvanced2)tinEdit;

            try
            {
                // Initialize New TIN
                IGeoDataset gds = inputFeatureClass as IGeoDataset;
                tinEdit.InitNew(gds.Extent);

                // Add Mass Points to TIN
                tinEdit.StartEditing();
                tinEdit.AddFromFeatureClass(inputFeatureClass, spatialFilter, uniqueField, uniqueField, esriTinSurfaceType.esriTinMassPoint);
                tinEdit.Refresh();

                // Get TIN Nodes
                ITinNodeCollection tinNodeCollection = (ITinNodeCollection)tinEdit;

                // Report Node Count
                message.AddMessage("Input Node Count: " + inputFeatureClass.FeatureCount(null).ToString());
                message.AddMessage("TIN Node Count: " + tinNodeCollection.NodeCount.ToString());

                // Open Output Feature Class
                IFeatureClass outputFeatureClass = m_GPUtilities.OpenFeatureClassFromString(outputParameterValue.GetAsText());

                // Debug Message
                message.AddMessage("Generating Polygons...");

                // Create Voronoi Polygons
                tinNodeCollection.ConvertToVoronoiRegions(outputFeatureClass, null, clippingPolygon, "", "");

                // Release COM Objects
                tinEdit.StopEditing(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tinNodeCollection);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tinEdit);
            }
            catch (Exception ex)
            {
                message.AddError(2, ex.Message);
            }
        }

        // This is the function name object for the Geoprocessing Function Tool. 
        // This name object is created and returned by the Function Factory.
        // The Function Factory must first be created before implementing this property.
        public IName FullName
        {
            get
            {
                IGPFunctionFactory functionFactory = new ThiessenFunctionFactory();
                return (IName)functionFactory.GetFunctionName(m_ToolName);
            }
        }

        // This is used to set a custom renderer for the output of the Function Tool.
        public object GetRenderer(IGPParameter pParam)
        {
            return null;
        }

        // This is the unique context identifier in a [MAP] file (.h). 
        // ESRI Knowledge Base article #27680 provides more information about creating a [MAP] file. 
        public int HelpContext
        {
            get { return 0; }
        }

        // This is the path to a .chm file which is used to describe and explain the function and its operation. 
        public string HelpFile
        {
            get { return ""; }
        }

        // This is used to return whether the function tool is licensed to execute.
        public bool IsLicensed()
        {
            IAoInitialize aoi = new AoInitializeClass();
            ILicenseInformation licInfo = (ILicenseInformation)aoi;

            string licName = licInfo.GetLicenseProductName(aoi.InitializedProduct());

            if (licName == "ArcView")
            {
                return true;
            }
            else
                return false;
        }

        // This is the name of the (.xml) file containing the default metadata for this function tool. 
        // The metadata file is used to supply the parameter descriptions in the help panel in the dialog. 
        // If no (.chm) file is provided, the help is based on the metadata file. 
        // ESRI Knowledge Base article #27000 provides more information about creating a metadata file.
        public string MetadataFile
        {
            get { return ""; }
        }

        // By default, the Toolbox will create a dialog based upon the parameters returned 
        // by the ParameterInfo property.
        public UID DialogCLSID
        {
            // DO NOT USE. INTERNAL USE ONLY.
            get { throw new Exception("The method or operation is not implemented."); }
        }

        #endregion
    }

    // FUNCTION FACTORY
    [
    Guid("F82746D9-3919-4F14-9872-122A9BDB02AF"),
    ComVisible(true)
    ]
    public class ThiessenFunctionFactory : IGPFunctionFactory
    {
        #region "Component Category Registration"
        [ComRegisterFunction()]
        static void Reg(string regKey)
        {
            GPFunctionFactories.Register(regKey);
        }

        [ComUnregisterFunction()]
        static void Unreg(string regKey)
        {
            GPFunctionFactories.Unregister(regKey);
        }
        #endregion

        // Utility Function added to create the function names.
        private IGPFunctionName CreateGPFunctionNames(long index)
        {
            IGPFunctionName functionName = new GPFunctionNameClass();
            functionName.MinimumProduct = esriProductCode.esriProductCodeProfessional;
            IGPName name;

            switch (index)
            {
                case (0):
                    name = (IGPName)functionName;
                    name.Category = "Geoprocessing";
                    name.Description = "Create Clipped Thiessen Polygons.";
                    name.DisplayName = "Thiessen Polygons";
                    name.Name = "Thiessen";
                    name.Factory = (IGPFunctionFactory)this;
                    break;
            }

            return functionName;
        }

        // Implementation of the Function Factory
        #region IGPFunctionFactory Members

        // This is the name of the function factory. 
        // This is used when generating the Toolbox containing the function tools of the factory.
        public string Name
        {
            get { return "Exponent"; }
        }

        // This is the alias name of the factory.
        public string Alias
        {
            get { return "exponent"; }
        }

        // This is the class id of the factory. 
        public UID CLSID
        {
            get
            {
                UID id = new UIDClass();
                id.Value = this.GetType().GUID.ToString("B");
                return id;
            }
        }

        // This method will create and return a function object based upon the input name.
        public IGPFunction GetFunction(string Name)
        {
            switch (Name)
            {
                case ("Thiessen"):
                    IGPFunction gpFunction = new GPThiessen.ThiessenFunction();
                    return gpFunction;
            }

            return null;
        }

        // This method will create and return a function name object based upon the input name.
        public IGPName GetFunctionName(string Name)
        {
            IGPName gpName = new GPFunctionNameClass();

            switch (Name)
            {
                case ("Thiessen"):
                    return (IGPName)CreateGPFunctionNames(0);

            }
            return null;
        }

        // This method will create and return an enumeration of function names that the factory supports.
        public IEnumGPName GetFunctionNames()
        {
            IArray nameArray = new EnumGPNameClass();
            nameArray.Add(CreateGPFunctionNames(0));
            return (IEnumGPName)nameArray;
        }

        // This method will create and return an enumeration of GPEnvironment objects. 
        // If tools published by this function factory required new environment settings, 
        // then you would define the additional environment settings here. 
        // This would be similar to how parameters are defined. 
        public IEnumGPEnvironment GetFunctionEnvironments()
        {
            return null;
        }

        #endregion
    }
}
        