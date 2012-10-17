using System;
using System.Collections;
using System.Data;
using System.Data.OleDb;
using System.Data.OracleClient;
using System.IO;
using System.Windows.Forms;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using ADOX;
using System.Reflection;
using System.Drawing;
using System.Collections.Generic;
using System.Xml;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.GeoDatabaseUI;
using ESRI.ArcGIS.Editor;
using System.Runtime.InteropServices;

namespace RegGSS
{
    public partial class FrmCoordinateSampleBuilder : Form
    {
        public IApplication m_application;
        int formHeight = 0;
        private string currentKeyFeatureValue = string.Empty;
        private string currentFormKeyValue = string.Empty;
        private ArrayList OIDList = new ArrayList();

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRandomApp_Click_1(object sender, EventArgs e)
        {
            IMxDocument pMxDocument = (IMxDocument)m_application.Document;
            IActiveView pActiveView = pMxDocument.ActiveView;
            IGraphicsContainer pGraphicsContainer = pMxDocument.FocusMap as IGraphicsContainer;
            try
            {
                statusMessage.Text = string.Empty;

                //if user opts out of submit and leaves graphics, delete them prior to zooming 
                //into new random feature

                pGraphicsContainer.DeleteAllElements();
                pActiveView.Refresh();
                pMxDocument.UpdateContents();

                //set the global variable --> key field 
                RegGSS.ClsGlobalVariables.CSBkeyField = cboKeyField.Text;

                //if neither a layername and county have been selected, give message, quit routine
                if (cboControlList.Text == "" && cboTestList.Text == "")
                {
                    statusMessage.Text = "Please select control and test layers from the above combo boxes.";
                    statusMessage.ForeColor = Color.Red;
                    return;
                }

                //if either a controlFileName or testFileName have been selected, give message, quit routine
                if (cboControlList.Text == "" || cboTestList.Text == "" || cboKeyField.Text == "")
                {
                    statusMessage.Text = "Please select a control, test and key field from the above combo boxes.";
                    statusMessage.ForeColor = Color.Red;
                    return;
                }

                CreateOIDList(cboTestList.Text);
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "btnRandomApp_Click_1", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
            }
            finally 
            {
                if (pMxDocument != null) { Marshal.ReleaseComObject(pMxDocument); pMxDocument = null; }
                if (pActiveView != null) { Marshal.ReleaseComObject(pActiveView); pActiveView = null; }
                if (pGraphicsContainer != null) { Marshal.ReleaseComObject(pGraphicsContainer); pGraphicsContainer = null; }     
            }
        }

        /// <summary>
        /// Determine how many graphics are selected.  Error message if not one test and one control graphic
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSubmit_Click(object sender, EventArgs e)
        {
            IMxDocument pMxDocument = (IMxDocument)m_application.Document;
            IActiveView pActiveView = pMxDocument.ActiveView;
            IGraphicsContainerSelect pGraphicsContainerSelect = pMxDocument.ActiveView as IGraphicsContainerSelect;
            IElement pElement = null;
            IMarkerElement pMarkerElement = null;
            Double markerSize = 0;
            try
            {
                statusMessage.Text = string.Empty;

                //Reduce height of form when clearing two-line StatusMessage.
                if (this.ClientSize.Height > formHeight)
                { this.ClientSize = new Size(this.ClientSize.Width, formHeight); }

                int idCount = 0;
                int idFound = 0;
                int controlFound = 0;
                double xcoord1 = 0;
                double xcoord2 = 0;
                double ycoord1 = 0;
                double ycoord2 = 0;

                if (pGraphicsContainerSelect.ElementSelectionCount == 0)
                {
                    statusMessage.Text = "No control points selected.  Nothing to submit.";
                    return;
                }
                if (pGraphicsContainerSelect.ElementSelectionCount == 1)
                {
                    statusMessage.Text = "Only one graphic is currently selected." + "\r\n" + "Please select one of each:  test and control graphic.";
                    statusMessage.ForeColor = Color.Red;

                    //Increase height of form to fit two-line StatusMessage.
                    this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
                }
                if (pGraphicsContainerSelect.ElementSelectionCount == 2)
                {
                    for (int index = 0; index < pGraphicsContainerSelect.ElementSelectionCount; index++)
                    {
                        pElement = pGraphicsContainerSelect.SelectedElement(index) as IElement;
                        if (pElement is IMarkerElement)
                        {
                            pMarkerElement = pElement as IMarkerElement;
                            markerSize = pMarkerElement.Symbol.Size;
                            if (markerSize == 15)
                            {
                                xcoord1 = pElement.Geometry.Envelope.XMin;
                                ycoord1 = pElement.Geometry.Envelope.YMin;
                                idFound += 1;
                            }
                            else
                            {
                                xcoord2 = pElement.Geometry.Envelope.XMin;
                                ycoord2 = pElement.Geometry.Envelope.YMin;
                                controlFound += 1;
                            }
                        }
                    }

                    if (idFound == 2)
                    {
                        statusMessage.Text = "More than one test graphic is selected. Please select only one.";
                        statusMessage.ForeColor = Color.Red;
                        return;
                    }
                    if (controlFound == 2)
                    {
                        statusMessage.Text = "More than one control graphic selected. Please select only one.";
                        statusMessage.ForeColor = Color.Red;
                        return;
                    }

                    populateCSBValues(xcoord1, xcoord2, ycoord1, ycoord2);
                    IGraphicsContainer pGraphicsContainer = pMxDocument.FocusMap as IGraphicsContainer;
                    pGraphicsContainer.DeleteAllElements();

                    pActiveView.Refresh();
                    pMxDocument.UpdateContents();

                    idCount = CheckIfCurrentIDAlreadyEntered();

                    if (idCount == 1)
                    {
                        //remove temporary test layer if submit was successful
                        RegGSS.ClsAddDeleteData.RemoveLayer("Selected test boundary");
                    }
                    DataGridViewRefresh();
                    UpdateControlPointCounter();
                }
                if (pGraphicsContainerSelect.ElementSelectionCount > 2)
                {
                    statusMessage.Text = "More than two coordinate pair graphics selected, Please select two.";
                    statusMessage.ForeColor = Color.Red;
                }
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "btnSubmit_Click", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
            }
            finally
            {
                if (pMxDocument != null) { Marshal.ReleaseComObject(pMxDocument); pMxDocument = null; }
                if (pActiveView != null) { Marshal.ReleaseComObject(pActiveView); pActiveView = null; }
                if (pGraphicsContainerSelect != null) { Marshal.ReleaseComObject(pGraphicsContainerSelect); pGraphicsContainerSelect = null; }
                if (pElement != null) { Marshal.ReleaseComObject(pElement); pElement = null; }
                if (pMarkerElement != null) { Marshal.ReleaseComObject(pMarkerElement); pMarkerElement = null; }
            }
        }

        private void cboControlList_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                statusMessage.Text = string.Empty;

                DataGridViewRefresh();
                UpdateControlPointCounter();
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "cboControlList_SelectedIndexChanged", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
            }
        }

        private void cboField_SelectedIndexChanged(object sender, EventArgs e)
        {
            IMxDocument pMxDocument = (IMxDocument)m_application.Document;
            IActiveView pActiveView = pMxDocument.ActiveView;
            IMap pMap = pMxDocument.FocusMap;
            IEnumLayer pEnumLayer = null;
            ILayer pLayer = null;
            IFeatureLayer pFeatureLayer = new FeatureLayerClass();
            IGeoFeatureLayer pGeoFeatureLayer = null;
            IFields pFields = null;
            IField pField = null;
            IDisplayTable displayTable = null;
            ITable table = null;
            IQueryFilter queryFilter = new QueryFilterClass();
            ICursor theCursor = null;
            IRow theRow = null;
            try
            {
                statusMessage.Text = string.Empty;

                //clear the value combobox
                cboValue.Items.Clear();

                //get the name of the selected field
                string selectedField = cboField.Text;

                //loop through all the layers
                pEnumLayer = pMap.get_Layers(null, true);
                pEnumLayer.Reset();
                pLayer = pEnumLayer.Next();
                string theString = string.Empty;

                while (pLayer != null)
                {
                    //locate the selected layer
                    if (pLayer.Name == cboTestList.Text)
                    {
                        pFeatureLayer = pLayer as IFeatureLayer;
                        pGeoFeatureLayer = pFeatureLayer as IGeoFeatureLayer;
                        pFields = pGeoFeatureLayer.DisplayFeatureClass.Fields;

                        RegGSS.ClsGlobalVariables.CSBtestLayerName = pFeatureLayer.FeatureClass.AliasName;

                        for (int counter = 1; counter < pFields.FieldCount; counter++)
                        {
                            //populate the field combobox with a list of the fields in the selected layer
                            pField = pFields.get_Field(counter);
                            if (pField.Name == selectedField)
                            {
                                displayTable = pFeatureLayer as IDisplayTable;
                                table = displayTable.DisplayTable;

                                string fieldType = Convert.ToString(pField.Type);
                                queryFilter = new QueryFilter();

                                if (fieldType == "esriFieldTypeString")
                                {
                                    queryFilter.WhereClause = selectedField + " <> ' '";
                                }
                                else
                                {
                                    queryFilter.WhereClause = selectedField + " <> 0";
                                }

                                theCursor = (ICursor)table.Search(queryFilter, true);
                                theRow = theCursor.NextRow();

                                int field = table.Fields.FindField(selectedField);
                                while (theRow != null)
                                {
                                    string value = Convert.ToString(theRow.get_Value(field));
                                    if (cboValue.Items.Contains(value) == false)
                                    {
                                        cboValue.Items.Add(value);
                                    }
                                    theRow = theCursor.NextRow();
                                }       //while (row != null)
                            }       //if (pField.Name == selectedField)
                        }       //for (int counter = 1; counter < pFields.FieldCount; counter++)
                    }       //if (pLayer.Name == cboTestList.Text)
                    pLayer = pEnumLayer.Next();
                }       //while (pLayer != null)
            }       //try
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "cboField_SelectedIndexChanged", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
            }
            finally
            {
                if (pMxDocument != null) { Marshal.ReleaseComObject(pMxDocument); pMxDocument = null; }
                if (pActiveView != null) { Marshal.ReleaseComObject(pActiveView); pActiveView = null; }
                if (pMap != null) { Marshal.ReleaseComObject(pMap); pMap = null; }
                if (pEnumLayer != null) { Marshal.ReleaseComObject(pEnumLayer); pEnumLayer = null; }
                if (pLayer != null) { Marshal.ReleaseComObject(pLayer); pLayer = null; }
                if (pFeatureLayer != null) { Marshal.ReleaseComObject(pFeatureLayer); pFeatureLayer = null; }
                if (pGeoFeatureLayer != null) { Marshal.ReleaseComObject(pGeoFeatureLayer); pGeoFeatureLayer = null; }
                if (pFields != null) { Marshal.ReleaseComObject(pFields); pFields = null; }
                if (pField != null) { Marshal.ReleaseComObject(pField); pField = null; }
                if (displayTable != null) { Marshal.ReleaseComObject(displayTable); displayTable = null; }
                if (table != null) { Marshal.ReleaseComObject(table); table = null; }
                if (queryFilter != null) { Marshal.ReleaseComObject(queryFilter); queryFilter = null; }
                if (theCursor != null) { Marshal.ReleaseComObject(theCursor); theCursor = null; }
                if (theRow != null) { Marshal.ReleaseComObject(theRow); theRow = null; }
            }
        }

        private void cboKeyField_SelectedIndexChanged(object sender, EventArgs e)
        {
            IMxDocument pMxDocument = (IMxDocument)m_application.Document;
            IActiveView pActiveView = pMxDocument.ActiveView;
            IMap pMap = pMxDocument.FocusMap;
            IEnumLayer pEnumLayer = null;
            ILayer pLayer = null;

            try
            {
                statusMessage.Text = string.Empty;

                //loop through all the layers
                pEnumLayer = pMap.get_Layers(null, true);
                pEnumLayer.Reset();
                pLayer = pEnumLayer.Next();
                string theString = string.Empty;

                //get the name of the test layer
                string testLayer = cboTestList.Text;

                while (pLayer != null)
                {
                    if (!(pLayer is IGroupLayer))
                    {
                        //locate the selected layer
                        if (pLayer.Name == testLayer)
                        {
                            IFeatureLayer pFeatureLayer = pLayer as IFeatureLayer;
                            RegGSS.ClsGlobalVariables.CSBtestLayerName = pFeatureLayer.FeatureClass.AliasName;
                        }   //if (pLayer.Name == testLayer)
                    }
                    pLayer = pEnumLayer.Next();
                }   //while (pLayer != null)
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "cboKeyField_SelectedIndexChanged", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
            }
            finally
            {
                if (pMxDocument != null) { Marshal.ReleaseComObject(pMxDocument); pMxDocument = null; }
                if (pActiveView != null) { Marshal.ReleaseComObject(pActiveView); pActiveView = null; }
                if (pMap != null) { Marshal.ReleaseComObject(pMap); pMap = null; }
                if (pLayer != null) { Marshal.ReleaseComObject(pLayer); pLayer = null; }
                if (pEnumLayer != null) { Marshal.ReleaseComObject(pEnumLayer); pEnumLayer = null; }
            }
        }

        private void cboTestList_SelectedIndexChanged(object sender, EventArgs e)
        {
            IMxDocument pMxDocument = (IMxDocument)m_application.Document;
            IActiveView pActiveView = pMxDocument.ActiveView;
            IMap pMap = pMxDocument.FocusMap;
            IEnumLayer pEnumLayer = null;
            ILayer pLayer = null;
            IFeatureLayer pFeatureLayer = new FeatureLayerClass();
            IGeoFeatureLayer pGeoFeatureLayer = null;
            IFields pFields = null;
            IField pField = null;

            try
            {
                statusMessage.Text = string.Empty;

                cboValue.Items.Clear();     //clear the value combobox
                cboField.Items.Clear();     //clear the field combobox

                if (cboControlList.Text != cboTestList.Text)
                {
                    //loop through all the layers
                    pEnumLayer = pMap.get_Layers(null, true);
                    pEnumLayer.Reset();
                    pLayer = pEnumLayer.Next();

                    while (pLayer != null)
                    {
                        if (!(pLayer is IGroupLayer))
                        {
                            //locate the selected layer
                            if (pLayer.Name == cboTestList.Text)
                            {
                                pFeatureLayer = pLayer as IFeatureLayer;
                                pGeoFeatureLayer = pFeatureLayer as IGeoFeatureLayer;
                                pFields = pGeoFeatureLayer.DisplayFeatureClass.Fields;

                                for (int counter = 1; counter < pFields.FieldCount; counter++)
                                {
                                    //populate the field combobox with a list of the fields in the selected layer
                                    pField = pFields.get_Field(counter);
                                    cboField.Items.Add(pField.Name);
                                    cboKeyField.Items.Add(pField.Name);
                                }
                            }
                        }
                        pLayer = pEnumLayer.Next();
                    }   //while (pLayer != null)
                }
                else
                {
                    statusMessage.Text = "Test layer can not be the same as the control layer." + "\r\n" + "Please select a different test layer.";
                    statusMessage.ForeColor = Color.Red;

                    //Increase height of form to fit two-line StatusMessage.
                    this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
                }
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "cboTestList_SelectedIndexChanged", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem. " + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
            }
            finally
            {
                if (pMxDocument != null) { Marshal.ReleaseComObject(pMxDocument); pMxDocument = null; }
                if (pActiveView != null) { Marshal.ReleaseComObject(pActiveView); pActiveView = null; }
                if (pMap != null) { Marshal.ReleaseComObject(pMap); pMap = null; }
                if (pEnumLayer != null) { Marshal.ReleaseComObject(pEnumLayer); pEnumLayer = null; }
                if (pLayer != null) { Marshal.ReleaseComObject(pLayer); pLayer = null; }
                if (pFeatureLayer != null) { Marshal.ReleaseComObject(pFeatureLayer); pFeatureLayer = null; }
                if (pGeoFeatureLayer != null) { Marshal.ReleaseComObject(pGeoFeatureLayer); pGeoFeatureLayer = null; }
                if (pFields != null) { Marshal.ReleaseComObject(pFields); pFields = null; }
                if (pField != null) { Marshal.ReleaseComObject(pField); pField = null; }
            }
        }

        private void cboTestList_MouseDown(object sender, MouseEventArgs e)
        {
            IMxDocument pMxDocument = (IMxDocument)m_application.Document;
            IActiveView pActiveView = pMxDocument.ActiveView;
            IMap pMap = pMxDocument.FocusMap;

            try
            {
                statusMessage.Text = string.Empty;

                //Stop code if no layers are loaded
                if (pMap.LayerCount == 0)
                {
                    statusMessage.Text = "There are currently no layers loaded.";
                    statusMessage.ForeColor = Color.Red;
                    return;
                }
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "cboTestList_MouseDown", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
            }
            finally
            {
                if (pMxDocument != null) { Marshal.ReleaseComObject(pMxDocument); pMxDocument = null; }
                if (pActiveView != null) { Marshal.ReleaseComObject(pActiveView); pActiveView = null; }
                if (pMap != null) { Marshal.ReleaseComObject(pMap); pMap = null; }
            }
        }

        /// <summary>
        /// determine if the current OID has previously been entered in the database
        /// </summary>
        /// <returns></returns>
        private int CheckIfCurrentIDAlreadyEntered()
        {
            OleDbConnection dataConnection = new OleDbConnection();
            OleDbCommand dataCommand = new OleDbCommand();
            OleDbDataReader dataReader = null;
            try
            {
                statusMessage.Text = string.Empty;

                //get the path for the current database
                string dblocation = RegGSS.ClsGlobalVariables.CSBDatabase;

                //extract the file name from it
                string tableName = System.IO.Path.GetFileNameWithoutExtension(dblocation);

                //get the current OID for the randomly selected feature
                int currentID = RegGSS.ClsGlobalVariables.CSBOIDNumber;

                int recordCount = 0;
                dataConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                                  "data source = " + dblocation;
                dataConnection.Open();

                dataCommand.Connection = dataConnection;
                dataCommand.CommandText = "SELECT count(id) " +
                                          "FROM [" + tableName + "] WHERE id = " + "'" + currentID + "'";

                dataReader = dataCommand.ExecuteReader();
                dataReader.Read();

                if (dataReader.HasRows == true)
                {
                    recordCount = dataReader.GetInt32(0);
                    return recordCount;
                }

                return 0;
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "CheckIfCurrentIDAlreadyEntered", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
                return 0;
            }
            finally
            {
                if (dataConnection != null) { if (dataConnection.State == System.Data.ConnectionState.Open) { dataConnection.Close(); dataConnection.Dispose(); } }
                if (dataCommand != null) { dataCommand.Dispose(); }
                if (dataReader != null) { if (dataReader.IsClosed == false) { dataReader.Close(); dataReader.Dispose(); } }
            }
        }

        /// <summary>
        /// create database in location specified by the user
        /// </summary>
        /// <param name="dbLocation"></param>
        private void createDatabase(string dbLocation)
        {
            try
            {
                statusMessage.Text = string.Empty;

                string tableName = System.IO.Path.GetFileNameWithoutExtension(dbLocation);
                grpBoxCSB.Text = "Database:  " + tableName;

                //If Access database does not already exists, create it
                if (File.Exists(dbLocation + ".accdb") == false)
                {
                    Catalog catalog = new Catalog();
                    string tmpString = string.Empty;
                    string filename = dbLocation;
                    tmpString = "Provider=Microsoft.ACE.OLEDB.12.0;";
                    tmpString += "Data Source=" + filename + ";Jet OLEDB:Engine Type=5";
                    catalog.Create(tmpString);
                    ADOX.Table newTable = new ADOX.Table();
                    newTable.Name = tableName;

                    newTable.Columns.Append("id", DataTypeEnum.adVarWChar, 25);
                    newTable.Columns.Append("xcoord1", DataTypeEnum.adDouble, 25);
                    newTable.Columns.Append("xcoord2", DataTypeEnum.adDouble, 25);
                    newTable.Columns.Append("ycoord1", DataTypeEnum.adDouble, 25);
                    newTable.Columns.Append("ycoord2", DataTypeEnum.adDouble, 25);
                    catalog.Tables.Append(newTable);

                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(newTable);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(catalog.Tables);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(catalog.ActiveConnection);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(catalog);
                }
                else
                {
                    switch (MessageBox.Show("Microsoft Access database already exists with this name." + "\r\n" + "Would you like to create one with a different name?", "Warning", MessageBoxButtons.OKCancel))
                    {
                        case DialogResult.OK:
                            redefineNewDatabse();
                            break;
                        case DialogResult.Cancel:
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "createDatabase", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
            }
        }

        /// <summary>
        ///Purpose:  Create a temporary layer used to get location of vertices
        ///          shade pink with 30% transparency
        /// </summary>
        /// <param name="pFeatureLayer">featurelayer of the currently selected application</param>
        private void createLayerFromSelectionSet(IFeatureLayer pFeatureLayer)
        {
            IMxDocument pMxDocument = (IMxDocument)m_application.Document;
            IFeatureLayerDefinition pFeatureLayerDefinition = pFeatureLayer as IFeatureLayerDefinition;
            IFeatureLayer pNewFeatureLayer = new FeatureLayerClass();
            IGeoFeatureLayer pGeoFeatureLayer = null;
            IRgbColor pLineColor = new RgbColorClass();
            ISimpleLineSymbol pSimpleLineSymbol = new SimpleLineSymbolClass();
            IRgbColor pFillColor = new RgbColorClass();
            ISimpleFillSymbol pSimpleFillSymbol = new SimpleFillSymbolClass();
            ISimpleRenderer pSimpleRenderer = new SimpleRendererClass();
            ILayerEffects pLayerEffects = null;
            try
            {
                statusMessage.Text = string.Empty;

                pNewFeatureLayer = pFeatureLayerDefinition.CreateSelectionLayer("Selected test boundary", true, null, null);
                pMxDocument.FocusMap.AddLayer(pNewFeatureLayer);

                pGeoFeatureLayer = pNewFeatureLayer as IGeoFeatureLayer;

                pLineColor.Red = 255;
                pLineColor.Green = 190;
                pLineColor.Blue = 190;

                pSimpleLineSymbol.Color = pLineColor;
                pSimpleLineSymbol.Width = 1;
                pSimpleLineSymbol.Style = esriSimpleLineStyle.esriSLSSolid;

                pFillColor.Red = 255;
                pFillColor.Green = 190;
                pFillColor.Blue = 190;

                pSimpleFillSymbol.Color = pFillColor;
                pSimpleFillSymbol.Outline = pSimpleLineSymbol;
                pSimpleFillSymbol.Style = esriSimpleFillStyle.esriSFSSolid;

                pSimpleRenderer.Label = "Selected test boundary";
                pSimpleRenderer.Symbol = (ISymbol)pSimpleFillSymbol;

                pGeoFeatureLayer.Renderer = (IFeatureRenderer)pSimpleRenderer;

                pLayerEffects = pGeoFeatureLayer as ILayerEffects;
                pLayerEffects.Transparency = 30;
                pMxDocument.ActivatedView.Refresh();

                ILayer selectedAppLayer = pGeoFeatureLayer;

                //Reorder layers to have temp layer below parcel area
                moveSelectedAppBelowParcelLayer(selectedAppLayer);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (pMxDocument != null) { Marshal.ReleaseComObject(pMxDocument); pMxDocument = null; }
                if (pFeatureLayerDefinition != null) { Marshal.ReleaseComObject(pFeatureLayerDefinition); pFeatureLayerDefinition = null; }
                if (pNewFeatureLayer != null) { Marshal.ReleaseComObject(pNewFeatureLayer); pNewFeatureLayer = null; }
                if (pGeoFeatureLayer != null) { Marshal.ReleaseComObject(pGeoFeatureLayer); pGeoFeatureLayer = null; }
                if (pLineColor != null) { Marshal.ReleaseComObject(pLineColor); pLineColor = null; }
                if (pSimpleLineSymbol != null) { Marshal.ReleaseComObject(pSimpleLineSymbol); pSimpleLineSymbol = null; }
                if (pFillColor != null) { Marshal.ReleaseComObject(pFillColor); pFillColor = null; }
                if (pSimpleFillSymbol != null) { Marshal.ReleaseComObject(pSimpleFillSymbol); pSimpleFillSymbol = null; }
                if (pSimpleRenderer != null) { Marshal.ReleaseComObject(pSimpleRenderer); pSimpleRenderer = null; }
                if (pLayerEffects != null) { Marshal.ReleaseComObject(pLayerEffects); pLayerEffects = null; }
            }
        }

        /// <summary>
        /// select a random OID number, zoom to that application and call method to create a graphic
        /// </summary>
        /// <param name="layerFileName"></param>
        public void CreateOIDList(string layerFileName)
        {
            IMxDocument pMxDocument = (IMxDocument)m_application.Document;
            IActiveView pActiveView = pMxDocument.ActiveView;
            IMap pMap = pMxDocument.FocusMap;
            IEnumLayer enumLayer = null;
            ILayer pLayer = null;
            IFeatureLayer pFeatureLayer = new FeatureLayerClass();
            IFeature pFeature = null;
            IFeatureCursor pFeatureCursor = null;
            IFields layerFields = null;
            IFeatureSelection pFeatureSelection = null;
            IQueryFilter pQueryFilter = new QueryFilterClass();
            IEnumFeature pEnumFeature = null;
            IEnvelope pEnvelope = new EnvelopeClass();

            try
            {
                statusMessage.Text = string.Empty;

                //populate the datagrid with all records from the open table
                DataGridViewRefresh();

                //populate a label with the count of records processed so far
                UpdateControlPointCounter();

                //Remove the temporary layer created from the previous run 
                RegGSS.ClsAddDeleteData.RemoveLayer("Selected test boundary");

                //Before zooming into a specific application, turn off the parcel layer (for speed)
                TurnOffControlLayer();

                //Stop code if no layers are loaded
                if (pMap.LayerCount == 0)
                {
                    statusMessage.Text = "There are currently no layers loaded.";
                    return;
                }

                enumLayer = pMap.get_Layers(null, true);
                enumLayer.Reset();
                pLayer = enumLayer.Next();
                pFeature = new Feature() as IFeature;

                //Populate the arraylist with all OID's from the selected layer
                while (pLayer != null)
                {
                    if (!(pLayer is IGroupLayer))
                    {
                        if (pLayer.Name == layerFileName)
                        {
                            pFeatureLayer = pLayer as IFeatureLayer;
                            if (OIDList.Count == 0)
                            {
                                pFeatureCursor = pFeatureLayer.FeatureClass.Search(null, false);
                                pFeature = pFeatureCursor.NextFeature();

                                while (pFeature != null)
                                {
                                    OIDList.Add(pFeature.OID);
                                    pFeature = pFeatureCursor.NextFeature();
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                    }   //if (!(pLayer is IGroupLayer))
                    pLayer = enumLayer.Next();
                }

                // Find the Object ID field.
                string OIDFieldName = string.Empty;
                layerFields = pFeatureLayer.FeatureClass.Fields as IFields;
                for (int fieldIndex = 0; fieldIndex < layerFields.FieldCount; fieldIndex++)
                {
                    if (layerFields.get_Field(fieldIndex).Type == esriFieldType.esriFieldTypeOID)
                    {
                        OIDFieldName = layerFields.get_Field(fieldIndex).Name;
                    }
                }

                //Boolean set up to catch when the random OID matches the selected county.  
                //Stop 'While' when this occurs
                bool OIDCountyMatch = false;
                string fieldValue = string.Empty;

                while (OIDCountyMatch == false)
                {
                    //get random OID number
                    Random randomOID = new Random();
                    int OIDNum = 0;
                    OIDNum = randomOID.Next(OIDList.Count);
                    RegGSS.ClsGlobalVariables.CSBOIDNumber = OIDNum;
                   
                    pFeatureSelection = pFeatureLayer as IFeatureSelection;

                    pQueryFilter = new QueryFilter();
                    pQueryFilter.WhereClause = OIDFieldName + " = " + OIDNum;
                    pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, null, null);
                    pFeatureSelection.SelectFeatures(pQueryFilter, esriSelectionResultEnum.esriSelectionResultNew, false);
                    pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, null, null);

                    pEnumFeature = pMap.FeatureSelection as IEnumFeature;
                    pEnumFeature.Reset();
                    pFeature = pEnumFeature.Next();
                    pEnvelope = new EnvelopeClass();

                    while (pFeature != null)
                    {
                        //Get the key field value selected by the user
                        string currentKeyField = RegGSS.ClsGlobalVariables.CSBkeyField;  //APP_NO
                        currentKeyFeatureValue = Convert.ToString(pFeature.get_Value(pFeature.Fields.FindField(currentKeyField)));  //030829-15
                        string currentFormField = cboField.Text;
                        currentFormKeyValue = cboValue.Text;

                        int idCount = 0;
                        idCount = CheckIfCurrentIDAlreadyEntered();
                        if (idCount == 0)
                        {
                            if (cboField.Text != "" && cboValue.Text != "")
                            {

                                if (getValueForKeyField(pFeatureLayer, currentKeyField, currentKeyFeatureValue, currentFormField) == true)
                                { return; }
                                else { break; }
                            }
                            else
                            {
                                pEnvelope.Union(pFeature.Extent);
                                pEnvelope.Expand(1.1, 1.1, true);
                                pActiveView.Extent = pEnvelope;

                                //Create a temporary layer with the selected application
                                createLayerFromSelectionSet(pFeatureLayer);

                                //Turn off visibility of the original layer 
                                pFeatureLayer.Visible = false;

                                OIDCountyMatch = true;
                                pActiveView.Refresh();
                                pMxDocument.UpdateContents();
                                return;
                            }
                        }   //if (idCount == 0)
                        pFeature = pEnumFeature.Next();
                    }   //while (pFeature != null)
                }   //while (OIDCountyMatch == false)
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "CreateOIDList", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
            }
            finally
            {
                if (pMxDocument != null) { Marshal.ReleaseComObject(pMxDocument); pMxDocument = null; }
                if (pActiveView != null) { Marshal.ReleaseComObject(pActiveView); pActiveView = null; }
                if (pMap != null) { Marshal.ReleaseComObject(pMap); pMap = null; }
                if (enumLayer != null) { Marshal.ReleaseComObject(enumLayer); enumLayer = null; }
                if (pLayer != null) { Marshal.ReleaseComObject(pLayer); pLayer = null; }
                if (pFeatureLayer != null) { Marshal.ReleaseComObject(pFeatureLayer); pFeatureLayer = null; }
                if (pFeature != null) { Marshal.ReleaseComObject(pFeature); pFeature = null; }
                if (pFeatureCursor != null) { Marshal.ReleaseComObject(pFeatureCursor); pFeatureCursor = null; }
                if (layerFields != null) { Marshal.ReleaseComObject(layerFields); layerFields = null; }
                if (pFeatureSelection != null) { Marshal.ReleaseComObject(pFeatureSelection); pFeatureSelection = null; }
                if (pQueryFilter != null) { Marshal.ReleaseComObject(pQueryFilter); pQueryFilter = null; }
                if (pEnumFeature != null) { Marshal.ReleaseComObject(pEnumFeature); pEnumFeature = null; }
                if (pEnvelope != null) { Marshal.ReleaseComObject(pEnvelope); pEnvelope = null; }
            }
        }

        /// <summary>
        /// for selected value, list current control points on the datagrid
        /// </summary>
        private void DataGridViewRefresh()
        {
            OleDbConnection dataConnection = new OleDbConnection();
            OleDbCommand dataCommand = new OleDbCommand();
            OleDbDataReader dataReader = null;

            try
            {
                statusMessage.Text = string.Empty;

                //get the path for the current database
                string dblocation = RegGSS.ClsGlobalVariables.CSBDatabase;

                //extract the file name from it
                string tableName = System.IO.Path.GetFileNameWithoutExtension(dblocation);
                int recordCount = 0;

                if (this.cboTestList.Text != string.Empty)
                {
                    if (this.cboControlList.Text != string.Empty)
                    {
                        dataConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                                          "data source = " + dblocation;
                        dataConnection.Open();

                        dataCommand.Connection = dataConnection;
                        dataCommand.CommandText = "SELECT count(id) " +
                                                  "FROM [" + tableName + "]";

                        dataReader = dataCommand.ExecuteReader();

                        while (dataReader.Read())
                        {
                            recordCount = dataReader.GetInt32(0);
                            if (recordCount != 0)
                            {
                                string strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                    "data source = " + dblocation;
                                string strCommandString =
                                     "Select xcoord1, xcoord2, ycoord1, ycoord2 from [" + tableName + "]";

                                OleDbDataAdapter DataAdapter = new OleDbDataAdapter(strCommandString, strConnectionString);

                                DataSet DataSet = new DataSet();
                                DataAdapter.Fill(DataSet, tableName);
                                dgrdRMS.DataSource = DataSet.Tables[tableName].DefaultView;
                            }
                        }
                        dataReader.Close();
                        dataConnection.Close();
                    }   //if (this.cboControlList.Text != string.Empty)
                }   //if (this.cboTestList.Text != string.Empty)
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (dataConnection != null) { if (dataConnection.State == System.Data.ConnectionState.Open) { dataConnection.Close(); dataConnection.Dispose(); } }
                if (dataCommand != null) { dataCommand.Dispose(); }
                if (dataReader != null) { if (dataReader.IsClosed == false) { dataReader.Close(); dataReader.Dispose(); } }
            }
        }

        public FrmCoordinateSampleBuilder()
        {
            InitializeComponent();
            formHeight = this.ClientSize.Height;
        }

        private void FrmRootMeanSquare_Load(object sender, EventArgs e)
        {
            IMxDocument pMxDocument = (IMxDocument)m_application.Document;
            IActiveView pActiveView = pMxDocument.ActiveView;
            IMap pMap = pMxDocument.FocusMap;
            IEnumLayer pEnumLayer = null;
            ILayer pLayer = null;

            try
            {
                statusMessage.Text = string.Empty;

                //loop through all the layers
                pEnumLayer = pMap.get_Layers(null, true);
                pEnumLayer.Reset();
                pLayer = pEnumLayer.Next();

                ArrayList TOCList = new ArrayList();
                while (pLayer != null)
                {
                    if (!(pLayer is IGroupLayer))
                    {
                        cboControlList.Items.Add(pLayer.Name);
                        cboTestList.Items.Add(pLayer.Name);
                    }
                    pLayer = pEnumLayer.Next();
                }

                btnRandomApp.Enabled = false;
                btnSubmit.Enabled = false;
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "FrmRootMeanSquare_Load", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
            }
            finally
            {
                if (pMxDocument != null) { Marshal.ReleaseComObject(pMxDocument); pMxDocument = null; }
                if (pActiveView != null) { Marshal.ReleaseComObject(pActiveView); pActiveView = null; }
                if (pMap != null) { Marshal.ReleaseComObject(pMap); pMap = null; }
                if (pEnumLayer != null) { Marshal.ReleaseComObject(pEnumLayer); pEnumLayer = null; }
                if (pLayer != null) { Marshal.ReleaseComObject(pLayer); pLayer = null; }
            }
        }

        /// <summary>
        ///  find field value where currentKeyField matches currentKeyFeatureValue
        ///  i.e.  find name of county where application number equals 030829-15 
        /// </summary>
        /// <param name="pFeatureLayer"></param>
        /// <param name="currentKeyField"></param>
        /// <param name="currentKeyFeatureValue"></param>
        /// <param name="currentFormField"></param>
        /// <returns></returns>
        private bool getValueForKeyField(ILayer pFeatureLayer, string currentKeyField, string currentKeyFeatureValue, string currentFormField)
        {
            IMxDocument pMxDocument = (IMxDocument)m_application.Document;
            IActiveView pActiveView = pMxDocument.ActiveView;
            IMap pMap = pMxDocument.FocusMap;
            IFeature pFeature = new Feature() as IFeature;
            IEnumLayer pEnumLayer = null;
            ILayer pLayer = null;
            IQueryFilter queryFilter = new QueryFilterClass();
            IFeatureSelection pFeatureSelection = null;
            IEnumFeature pEnumFeature = null;
            IFeatureLayer featureLayer = new FeatureLayerClass();
            IDataset pDataSet = null;
            IEnvelope pEnvelope = new EnvelopeClass();

            try
            {
                statusMessage.Text = string.Empty;

                //get the dataset for the passed in layer
                IDataset testDataset = pFeatureLayer as IDataset;

                pEnumLayer = pMap.get_Layers(null, true);
                pEnumLayer.Reset();
                pLayer = pEnumLayer.Next();

                string sqlString = string.Empty;

                while (pLayer != null)
                {
                    if (!(pLayer is IGroupLayer))
                    {
                        featureLayer = pLayer as IFeatureLayer;
                        if (featureLayer != null)
                        {
                            pDataSet = featureLayer as IDataset;
                            if (pDataSet.BrowseName == testDataset.BrowseName)
                            {
                                sqlString = "" + currentKeyField + " = '" + currentKeyFeatureValue + "'";

                                queryFilter = new QueryFilterClass();
                                queryFilter.WhereClause = sqlString;

                                pFeatureSelection = pFeatureLayer as IFeatureSelection;

                                pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, null, null);
                                pFeatureSelection.SelectFeatures(queryFilter, esriSelectionResultEnum.esriSelectionResultNew, false);
                                pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, null, null);

                                pEnumFeature = pMap.FeatureSelection as IEnumFeature;
                                pEnumFeature.Reset();
                                pFeature = pEnumFeature.Next();

                                string currentFeatureField = string.Empty;
                                currentFeatureField = Convert.ToString(pFeature.get_Value(pFeature.Fields.FindField(currentFormField)));

                                //Continue if sub-sample value matches the field value from random pfeature
                                if (currentFeatureField == cboValue.Text)
                                {
                                    pEnvelope.Union(pFeature.Extent);
                                    pEnvelope.Expand(1.1, 1.1, true);
                                    pActiveView.Extent = pEnvelope;

                                    //Create a temporary layer with the selected application
                                    createLayerFromSelectionSet(featureLayer);

                                    //Turn off visibility of the original layer 
                                    pFeatureLayer.Visible = false;

                                    pActiveView.Refresh();
                                    pMxDocument.UpdateContents();
                                    return true;
                                }
                                else
                                {
                                    return false;
                                }
                            }
                        }
                    }
                    pLayer = pEnumLayer.Next();
                }
                return false;
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "getValueForKeyField", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
                return false;
            }
            finally
            {
                if (pMxDocument != null) { Marshal.ReleaseComObject(pMxDocument); pMxDocument = null; }
                if (pActiveView != null) { Marshal.ReleaseComObject(pActiveView); pActiveView = null; }
                if (pMap != null) { Marshal.ReleaseComObject(pMap); pMap = null; }
                if (pFeature != null) { Marshal.ReleaseComObject(pFeature); pFeature = null; }
                if (queryFilter != null) { Marshal.ReleaseComObject(queryFilter); queryFilter = null; }
                if (pFeatureSelection != null) { Marshal.ReleaseComObject(pFeatureSelection); pFeatureSelection = null; }
                if (pEnumFeature != null) { Marshal.ReleaseComObject(pEnumFeature); pEnumFeature = null; }
            }
        }

        /// <summary>
        /// Join a table to a feature class
        /// </summary>
        /// <param name="pFeatureLayer"></param>
        /// <param name="pTable"></param>
        /// <param name="layerField"></param>
        /// <param name="tableField"></param>
        /// <returns></returns>
        internal bool JoinTableToLayer(IGeoFeatureLayer pFeatureLayer, ITable pTable,
                                   string layerField, string tableField)
        {
            IMemoryRelationshipClassFactory pMemoryRelationshipClassFactory = new MemoryRelationshipClassFactoryClass();
            IRelationshipClass pRelationshipClass = null;
            IDisplayRelationshipClass pDisplayRelationshipClass = null;

            try
            {
                statusMessage.Text = string.Empty;

                //set the relationship one to one
                pRelationshipClass = pMemoryRelationshipClassFactory.Open("Join", (IObjectClass)pFeatureLayer.DisplayFeatureClass, layerField, (IObjectClass)pTable, tableField, "forward", "backward", esriRelCardinality.esriRelCardinalityOneToOne);

                //perform the join
                pDisplayRelationshipClass = pFeatureLayer as IDisplayRelationshipClass;
                pDisplayRelationshipClass.DisplayRelationshipClass(pRelationshipClass, esriJoinType.esriLeftOuterJoin);
                return true;
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "JoinTableToLayer", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
                return false;
            }
            finally
            {
                if (pMemoryRelationshipClassFactory != null) { Marshal.ReleaseComObject(pMemoryRelationshipClassFactory); pMemoryRelationshipClassFactory = null; }
                if (pRelationshipClass != null) { Marshal.ReleaseComObject(pRelationshipClass); pRelationshipClass = null; }
                if (pDisplayRelationshipClass != null) { Marshal.ReleaseComObject(pDisplayRelationshipClass); pDisplayRelationshipClass = null; }            }
        }

        /// <summary>
        /// reorder the temporary layer (Selected App) below parcel layer to better visualize errors
        /// </summary>
        /// <param name="selectedAppLayer"></param>
        private void moveSelectedAppBelowParcelLayer(ILayer selectedAppLayer)
        {
            IMxDocument pMxDocument = (IMxDocument)m_application.Document;
            IActiveView pActiveView = pMxDocument.ActiveView;
            IMap pMap = pMxDocument.FocusMap;
            IEnumLayer enumLayer = null;
            ILayer pLayer = null;
            IFeatureLayer pFeatureLayer = new FeatureLayerClass();

            try
            {
                statusMessage.Text = string.Empty;

                enumLayer = pMap.get_Layers(null, true);
                enumLayer.Reset();
                pLayer = enumLayer.Next();

                //loop through all layers, find parcel layer, move selected app layer below it
                for (int indexNo = 0; indexNo < pMap.LayerCount; indexNo++)
                {
                    if (!(pLayer is IGroupLayer))
                    {
                        if (pLayer.Name == cboControlList.Text)
                        {
                            pActiveView.FocusMap.MoveLayer(selectedAppLayer, indexNo + 1);
                            pActiveView.Refresh();
                            pMxDocument.UpdateContents();
                            break;
                        }
                    }
                    pLayer = enumLayer.Next();
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (pMxDocument != null) { Marshal.ReleaseComObject(pMxDocument); pMxDocument = null; }
                if (pActiveView != null) { Marshal.ReleaseComObject(pActiveView); pActiveView = null; }
                if (pMap != null) { Marshal.ReleaseComObject(pMap); pMap = null; }
                if (pFeatureLayer != null) { Marshal.ReleaseComObject(pFeatureLayer); pFeatureLayer = null; }
                if (pLayer != null) { Marshal.ReleaseComObject(pLayer); pLayer = null; }
            }
        }

        /// <summary>
        /// get location where user wants to store the new database
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void newDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                statusMessage.Text = string.Empty;

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.AddExtension = true;
                saveFileDialog.DefaultExt = ".accdb";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //call method to create the database at the specified location
                    string dbLocation = saveFileDialog.FileName;
                    createDatabase(dbLocation);
                    RegGSS.ClsGlobalVariables.CSBDatabase = dbLocation;

                    btnRandomApp.Enabled = true;
                    btnSubmit.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "newDatabaseToolStripMenuItem_Click", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
            }
        }

        private void openDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OleDbConnection dataConnection = new OleDbConnection();
            OleDbCommand dataCommand = new OleDbCommand();
            OleDbDataReader dataReader = null;

            try
            {
                statusMessage.Text = string.Empty;

                //let user select a database to open, populate datagrid with values previously created
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = true;
                openFileDialog.FilterIndex = 0;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (string name in openFileDialog.FileNames)
                    {
                        string dbLocation = openFileDialog.FileName;
                        RegGSS.ClsGlobalVariables.CSBDatabase = dbLocation;
                        string tableName = System.IO.Path.GetFileNameWithoutExtension(dbLocation);
                        grpBoxCSB.Text = "Database:  " + tableName;

                        btnRandomApp.Enabled = true;
                        btnSubmit.Enabled = true;
                        dataConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                                          "data source = " + dbLocation;
                        dataConnection.Open();

                        dataCommand.Connection = dataConnection;
                        dataCommand.CommandText = "SELECT count(id) FROM [" + tableName + "]";
                        dataReader = dataCommand.ExecuteReader();

                        while (dataReader.Read())
                        {
                            if (dataReader.GetInt32(0) != 0)
                            {
                                string strConnectionString =
                                    "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                    "data source = " + dbLocation;
                                string strCommandString =
                                     "Select xcoord1, xcoord2, ycoord1, ycoord2 from [" + tableName + "]";

                                OleDbDataAdapter DataAdapter = new OleDbDataAdapter(strCommandString, strConnectionString);

                                DataSet DataSet = new DataSet();
                                DataAdapter.Fill(DataSet, tableName);
                                dgrdRMS.DataSource = DataSet.Tables[tableName].DefaultView;
                            }
                        }
                        dataReader.Close();
                        dataConnection.Close();
                    }   //foreach (string name in openFileDialog.FileNames)
                }   //if (openFileDialog.ShowDialog() == DialogResult.OK)
            } //try
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "openDatabaseToolStripMenuItem_Click", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
            }
            finally
            {
                if (dataConnection != null) { if (dataConnection.State == System.Data.ConnectionState.Open) { dataConnection.Close(); dataConnection.Dispose(); } }
                if (dataCommand != null) { dataCommand.Dispose(); }
                if (dataReader != null) { if (dataReader.IsClosed == false) { dataReader.Close(); dataReader.Dispose(); } }
            }
        }

        /// <summary>
        /// add id number and x / y coordinates to the access database
        /// </summary>
        /// <param name="xcoord1">x coordinate for vertice of test boundary</param>
        /// <param name="xcoord2">y coordinate for vertice of test boundary</param>
        /// <param name="ycoord1">x coordinate for vertice of parcel boundary</param>
        /// <param name="ycoord2">y coordinate for vertice of parcel boundary</param>
        private void populateCSBValues(double xcoord1, double xcoord2, double ycoord1, double ycoord2)
        {
            OleDbConnection dataConnection = new OleDbConnection();
            OleDbCommand dataCommand = new OleDbCommand();
            OleDbDataReader dataReader = null;
            try
            {
                statusMessage.Text = string.Empty;

                //get the path for the current database
                string dblocation = RegGSS.ClsGlobalVariables.CSBDatabase;

                //extract the file name from it
                string tableName = System.IO.Path.GetFileNameWithoutExtension(dblocation);

                //get the current object Id for the randomly selected feature
                int currentID = RegGSS.ClsGlobalVariables.CSBOIDNumber;

                dataConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                     "Data Source = " + dblocation;
                dataConnection.Open();

                dataCommand.Connection = dataConnection;
                dataCommand.CommandText = "INSERT INTO [" + tableName + "] (id, xcoord1, xcoord2, ycoord1, ycoord2) " +
                                          "VALUES ('" + currentID + "','" + xcoord1 + "','" + xcoord2 + "','" + ycoord1 + "','" + ycoord2 + "')";
                dataReader = dataCommand.ExecuteReader();
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (dataConnection != null) { if (dataConnection.State == System.Data.ConnectionState.Open) { dataConnection.Close(); dataConnection.Dispose(); } }
                if (dataCommand != null) { dataCommand.Dispose(); }
                if (dataReader != null) { if (dataReader.IsClosed == false) { dataReader.Close(); dataReader.Dispose(); } }
            }
        }

        /// <summary>
        /// get location where user wants to store the new database
        /// </summary>
        private void redefineNewDatabse()
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "*.accdb|AllFiles(*.*)";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //call method to create the database at the specified location
                    string dbLocation = saveFileDialog.FileName;
                    createDatabase(dbLocation);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// set visibility of control layer to off. 
        /// Prevent slow refresh time if app is zoomed to a large extent
        /// </summary>
        private void TurnOffControlLayer()
        {
            IMxDocument pMxDocument = (IMxDocument)m_application.Document;
            IActiveView pActiveView = pMxDocument.ActiveView;
            IMap pMap = pMxDocument.FocusMap;
            IEnumLayer enumLayer = null;
            ILayer pLayer = null;
            IFeatureLayer pFeatureLayer = new FeatureLayerClass();

            try
            {
                statusMessage.Text = string.Empty;

                //Stop code if no layers are loaded
                if (pMap.LayerCount == 0)
                {
                    return;
                }

                enumLayer = pMap.get_Layers(null, true);
                enumLayer.Reset();
                pLayer = enumLayer.Next();

                while (pLayer != null)
                {
                    if (!(pLayer is IGroupLayer))
                    {
                        if (pLayer.Name == cboControlList.Text)
                        {
                            pFeatureLayer = pLayer as IFeatureLayer;
                            pFeatureLayer.Selectable = true;

                            if (chkKeepVisible.Checked == true)
                            {
                                pFeatureLayer.Visible = true;
                                break;
                            }
                            else
                            {
                                pFeatureLayer.Visible = false;
                                break;
                            }
                        }
                    }
                    pLayer = enumLayer.Next();
                }
                pMxDocument.ActiveView.Refresh();
                pMxDocument.UpdateContents();
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "TurnOffControlLayer", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
            }
            finally
            {
                if (pMxDocument != null) { Marshal.ReleaseComObject(pMxDocument); pMxDocument = null; }
                if (pActiveView != null) { Marshal.ReleaseComObject(pActiveView); pActiveView = null; }
                if (pMap != null) { Marshal.ReleaseComObject(pMap); pMap = null; }
                if (pFeatureLayer != null) { Marshal.ReleaseComObject(pFeatureLayer); pFeatureLayer = null; }
                if (pLayer != null) { Marshal.ReleaseComObject(pLayer); pLayer = null; }
                if (enumLayer != null) { Marshal.ReleaseComObject(enumLayer); enumLayer = null; }
            }
        }

        /// <summary>
        /// Update counter of how many control points have been processed so far
        /// </summary>
        private void UpdateControlPointCounter()
        {
            OleDbConnection dataConnection = new OleDbConnection();
            OleDbCommand dataCommand = new OleDbCommand();
            OleDbDataReader dataReader = null;

            try
            {
                statusMessage.Text = string.Empty;

                if (this.cboTestList.Text != string.Empty)
                {
                    if (this.cboControlList.Text != string.Empty)
                    {
                        int recordCount = 0;
                       
                        //get the path for the current database
                        string dblocation = RegGSS.ClsGlobalVariables.CSBDatabase;
                       
                        //extract the file name from it
                        string tableName = System.IO.Path.GetFileNameWithoutExtension(dblocation);

                        dataConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                                          "data source = " + dblocation;
                        dataConnection.Open();
                        dataCommand.Connection = dataConnection;
                        dataCommand.CommandText = "SELECT count(id) " +
                                                  "FROM [" + tableName + "]";

                        dataReader = dataCommand.ExecuteReader();
                        while (dataReader.Read())
                        {
                            recordCount = dataReader.GetInt32(0);
                            if (recordCount != 0)
                            {
                                lblControlPoints.Text = "Number of coordinate pairs:" + Convert.ToString(recordCount);
                            }
                            else
                            {
                                this.lblRecordCounter.Text = Convert.ToString(0);
                            }
                        }
                    }   //if (this.cboControlList.Text != string.Empty)
                }   //if (this.cboTestList.Text != string.Empty)
            }
            catch (Exception ex)
            {
                ClsLogErrors.LogError(ex.StackTrace, ex.Message, "UpdateControlPointCounter", "FrmCoordinateSampleBuilder");
                statusMessage.Text = "The RegGSS Extension encountered a problem." + "\r\n" + "The Regulatory GIS Section has been notified.";
                statusMessage.ForeColor = Color.Black;

                //Increase height of form to fit two-line StatusMessage.
                this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + 10);
            }
            finally
            {
                if (dataConnection != null) { if (dataConnection.State == System.Data.ConnectionState.Open) { dataConnection.Close(); dataConnection.Dispose(); } }
                if (dataCommand != null) { dataCommand.Dispose(); }
                if (dataReader != null) { if (dataReader.IsClosed == false) { dataReader.Close(); dataReader.Dispose(); } }
            }
        }

    }
}