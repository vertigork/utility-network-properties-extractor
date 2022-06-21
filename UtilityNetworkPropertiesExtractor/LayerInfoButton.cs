/*
   Copyright 2021 Esri
   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at
       http://www.apache.org/licenses/LICENSE-2.0
   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS, 
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/
using ArcGIS.Core.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Dialogs;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace UtilityNetworkPropertiesExtractor
{
    internal class LayerInfoButton : Button
    {
        private static string _fileName = String.Empty;

        protected async override void OnClick()
        {
            try
            {
                await ExtractLayerInfoAsync();
                MessageBox.Show("Directory: " + Common.ExtractFilePath + Environment.NewLine + "File Name: " + _fileName, "CSV file has been generated");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Extract Layer Info");
            }
        }

        public static Task ExtractLayerInfoAsync()
        {
            return QueuedTask.Run(() =>
            {
                Common.CreateOutputDirectory();

                string dateFormatted = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string outputLayerFileName = Path.Combine(Common.ExtractFilePath, string.Format("{0}_{1}_LayerInfo.csv", dateFormatted, Common.GetProProjectName()));
                _fileName = outputLayerFileName;
                string outputLabelFileName = Path.Combine(Common.ExtractFilePath, string.Format("{0}_{1}_LabelInfo.csv", dateFormatted, Common.GetProProjectName()));
                string outputPopupFileName = Path.Combine(Common.ExtractFilePath, string.Format("{0}_{1}_PopupInfo.csv", dateFormatted, Common.GetProProjectName()));
                string outputDefQeryFileName = Path.Combine(Common.ExtractFilePath, string.Format("{0}_{1}_DefQueryInfo.csv", dateFormatted, Common.GetProProjectName()));

                int layerPos = 1;
                string prevGroupLayerName = string.Empty;
                string layerContainer = string.Empty;
                bool increaseLayerPos = false;

                var layerInfoList = new List<CsvLayerInfo>();
                var popupInfoList = new List<CsvPopupInfo>();
                var labelInfoList = new List<CsvLabelInfo>();
                var definitionQueryInfoList = new List<CsvDefinitionQueryInfo>();

                List<Layer> layerList = MapView.Active.Map.GetLayersAsFlattenedList().OfType<Layer>().ToList();
                foreach (Layer layer in layerList)
                {
                    try
                    {
                        layerContainer = layer.Parent.ToString();
                        if (layerContainer != MapView.Active.Map.Name) // Group layer
                        {
                            if (layerContainer != prevGroupLayerName)
                                prevGroupLayerName = layerContainer;
                        }
                        else
                            layerContainer = string.Empty;

                        if (layer is FeatureLayer featureLayer)
                        {
                            CIMFeatureLayer cimFeatureLayerDef = layer.GetDefinition() as CIMFeatureLayer;
                            CIMFeatureTable cimFeatureTable = cimFeatureLayerDef.FeatureTable;
                            CIMExpressionInfo cimExpressionInfo = cimFeatureTable.DisplayExpressionInfo; ;

                            //Primary Display Field
                            string displayField = cimFeatureTable.DisplayField;
                            if (cimExpressionInfo != null)
                                displayField = cimExpressionInfo.Expression.Replace("\"", "'");  //double quotes messes up the delimeters in the CSV

                            //Labeling
                            string labelExpression = string.Empty;
                            string labelMinScale = string.Empty;
                            string labelMaxScale = string.Empty;
                            if (cimFeatureLayerDef.LabelClasses != null)
                            {
                                if (cimFeatureLayerDef.LabelClasses.Length != 0)
                                {
                                    List<CIMLabelClass> cimLabelClassList = cimFeatureLayerDef.LabelClasses.ToList();
                                        
                                    CIMLabelClass cimLabelClass = cimLabelClassList.FirstOrDefault();
                                    labelExpression = cimLabelClass.Expression.Replace("\"", "'").Replace("\n", " ");  //double quotes messes up the delimeters in the CSV
                                    labelMinScale = GetScaleValue(cimLabelClass.MinimumScale);
                                    labelMaxScale = GetScaleValue(cimLabelClass.MaximumScale);

                                    foreach (var labelClass in cimLabelClassList)
                                    {
                                        var labelInfo = new CsvLabelInfo
                                        {
                                            LayerPos = layerPos.ToString(),
                                            LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                            LayerType = "Feature Layer",
                                            GroupLayerName = Common.EncloseStringInDoubleQuotes(layerContainer),
                                            IsLabelVisible = cimLabelClass.Visibility.ToString(),
                                            LabelClassName = cimLabelClass.Name,
                                            LabelExpression = cimLabelClass.Expression.Replace("\"", "'").Replace("\n", " "),
                                            LabelMinScale = GetScaleValue(cimLabelClass.MinimumScale),
                                            LabelMaxScale = GetScaleValue(cimLabelClass.MaximumScale),
                                        };
                                        labelInfoList.Add(labelInfo);
                                    }
                                }
                            }

                            //symbology
                            DetermineSymbology(cimFeatureLayerDef, out string primarySymbology, out string field1, out string field2, out string field3);

                            string subtypeValue = string.Empty;
                            if (featureLayer.IsSubtypeLayer)
                                subtypeValue = featureLayer.SubtypeValue.ToString();

                            var currentTable = featureLayer.GetTable();
                            CsvLayerInfo rec = new CsvLayerInfo()
                            {
                                LayerPos = layerPos.ToString(),
                                LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                LayerType = "Feature Layer",
                                GroupLayerName = Common.EncloseStringInDoubleQuotes(layerContainer),
                                IsVisible = layer.IsVisible.ToString(),
                                LayerSource = currentTable != null ? currentTable.GetPath().ToString() : String.Empty,
                                ClassName = currentTable != null ? currentTable.GetName() : String.Empty,
                                IsSubtypeLayer = featureLayer.IsSubtypeLayer.ToString(),
                                SubtypeValue = subtypeValue,
                                GeometryType = featureLayer.ShapeType.ToString(),
                                IsSelectable = featureLayer.IsSelectable.ToString(),
                                IsSnappable = featureLayer.IsSnappable.ToString(),
                                IsEditable = featureLayer.IsEditable.ToString(),
                                RefreshRate = cimFeatureLayerDef.RefreshRate.ToString(),
                                DefinitionQueryName = Common.EncloseStringInDoubleQuotes(featureLayer.DefinitionFilter.Name),
                                DefinitionQuery = Common.EncloseStringInDoubleQuotes(featureLayer.DefinitionFilter.DefinitionExpression),
                                MinScale = GetScaleValue(layer.MinScale),
                                MaxScale = GetScaleValue(layer.MaxScale),
                                ShowMapTips = cimFeatureLayerDef.ShowMapTips.ToString(),
                                PrimarySymbology = primarySymbology,
                                SymbologyField1 = field1,
                                SymbologyField2 = field2,
                                SymbologyField3 = field3,
                                EditTemplateCount = cimFeatureLayerDef.FeatureTemplates?.Count().ToString(),
                                DisplayField = Common.EncloseStringInDoubleQuotes(displayField),
                                IsLabelVisible = featureLayer.IsLabelVisible.ToString(),
                                LabelExpression = Common.EncloseStringInDoubleQuotes(labelExpression),
                                LabelMinScale = labelMinScale,
                                LabelMaxScale = labelMaxScale
                            };
                            layerInfoList.Add(rec);
                            increaseLayerPos = true;

                            foreach (var definitionFilter in featureLayer.GetDefinitionFilters())
                            {
                                var definitionQueryInfo = new CsvDefinitionQueryInfo
                                {
                                    LayerPos = layerPos.ToString(),
                                    LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                    LayerType = "Feature Layer",
                                    GroupLayerName = Common.EncloseStringInDoubleQuotes(layerContainer),
                                    DefinitionQueryName = definitionFilter.Name,
                                    DefinitionQuery = Common.EncloseStringInDoubleQuotes(definitionFilter.DefinitionExpression),
                                };
                                definitionQueryInfoList.Add(definitionQueryInfo);
                            }

                            if (cimFeatureLayerDef.EnableDisplayFilters)
                            {
                                CIMDisplayFilter[] cimDisplayFilterChoices = cimFeatureLayerDef.DisplayFilterChoices;
                                CIMDisplayFilter[] cimDisplayFilter = cimFeatureLayerDef.DisplayFilters;
                                AddDisplayFiltersToList(rec, cimDisplayFilterChoices, cimDisplayFilter, ref layerInfoList);
                            }

                            //Include Pop-up expressions if exist
                            if (cimFeatureLayerDef.PopupInfo != null)
                            {
                                if (cimFeatureLayerDef.PopupInfo.ExpressionInfos != null)
                                {
                                    bool popupExprVisibility = false;
                                    for (int i = 0; i < cimFeatureLayerDef.PopupInfo.ExpressionInfos.Length; i++)
                                    {
                                        //determine if expression is visible in popup
                                        CIMMediaInfo[] cimMediaInfos = cimFeatureLayerDef.PopupInfo.MediaInfos;
                                        for (int j = 0; j < cimMediaInfos.Length; j++)
                                        {
                                            if (cimMediaInfos[j] is CIMTableMediaInfo cimTableMediaInfo)
                                            {
                                                string[] fields = cimTableMediaInfo.Fields;
                                                for (int k = 0; k < fields.Length; k++)
                                                {
                                                    if (fields[k] == "expression/" + cimFeatureLayerDef.PopupInfo.ExpressionInfos[i].Name)
                                                    {
                                                        popupExprVisibility = true;
                                                        break;
                                                    }
                                                }
                                            }

                                            //Break out of 2nd loop (j) if already found the expression
                                            if (popupExprVisibility)
                                                break;
                                        }

                                        //Write popup info
                                        var popupRec = new CsvPopupInfo
                                        {
                                            LayerPos = layerPos.ToString(),
                                            LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                            LayerType = "Feature Layer",
                                            GroupLayerName = Common.EncloseStringInDoubleQuotes(layerContainer),
                                            PopupExpresssionName = cimFeatureLayerDef.PopupInfo.ExpressionInfos[i].Name,
                                            PopupExpresssionTitle = Common.EncloseStringInDoubleQuotes(cimFeatureLayerDef.PopupInfo.ExpressionInfos[i].Title.Replace("\"", "'")),
                                            PopupExpresssionVisible = popupExprVisibility.ToString(),
                                            PopupExpressionArcade = Common.EncloseStringInDoubleQuotes(cimFeatureLayerDef.PopupInfo.ExpressionInfos[i].Expression.Replace("\"", "'"))
                                        };
                                        popupInfoList.Add(popupRec);
                                    }
                                }
                            }
                        }
                        else if (layer is SubtypeGroupLayer subtypeGroupLayer)
                        {
                            CIMSubtypeGroupLayer cimSubtypeGroupLayer = layer.GetDefinition() as CIMSubtypeGroupLayer;

                            CsvLayerInfo rec = new CsvLayerInfo()
                            {
                                LayerPos = layerPos.ToString(),
                                LayerType = "Subtype Group Layer",
                                GroupLayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                IsVisible = layer.IsVisible.ToString(),
                                DefinitionQuery = Common.EncloseStringInDoubleQuotes(subtypeGroupLayer.DefinitionFilter.DefinitionExpression),
                                MinScale = GetScaleValue(layer.MinScale),
                                MaxScale = GetScaleValue(layer.MaxScale)
                            };
                            layerInfoList.Add(rec);
                            increaseLayerPos = true;

                            foreach (var definitionFilter in subtypeGroupLayer.GetDefinitionFilters())
                            {
                                var definitionQueryInfo = new CsvDefinitionQueryInfo
                                {
                                    LayerPos = layerPos.ToString(),
                                    LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                    LayerType = "Subtype Group Layer",
                                    GroupLayerName = Common.EncloseStringInDoubleQuotes(layerContainer),
                                    DefinitionQueryName = definitionFilter.Name,
                                    DefinitionQuery = Common.EncloseStringInDoubleQuotes(definitionFilter.DefinitionExpression),
                                };
                                definitionQueryInfoList.Add(definitionQueryInfo);
                            }

                            if (cimSubtypeGroupLayer.EnableDisplayFilters)
                            {
                                CIMDisplayFilter[] cimDisplayFilterChoices = cimSubtypeGroupLayer.DisplayFilterChoices;
                                CIMDisplayFilter[] cimDisplayFilter = cimSubtypeGroupLayer.DisplayFilters;
                                AddDisplayFiltersToList(rec, cimDisplayFilterChoices, cimDisplayFilter, ref layerInfoList);
                            }
                        }
                        else if (layer is GroupLayer groupLayer)
                        {
                            CsvLayerInfo rec = new CsvLayerInfo()
                            {
                                LayerPos = layerPos.ToString(),
                                GroupLayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                LayerType = "Group Layer",
                                IsVisible = layer.IsVisible.ToString(),
                                MinScale = GetScaleValue(layer.MinScale),
                                MaxScale = GetScaleValue(layer.MaxScale)
                            };
                            layerInfoList.Add(rec);
                            increaseLayerPos = true;
                        }
                        else if (layer is AnnotationLayer annotationLayer)
                        {
                            CIMAnnotationLayer cimAnnotationLayer = layer.GetDefinition() as CIMAnnotationLayer;

                            CsvLayerInfo rec = new CsvLayerInfo()
                            {
                                LayerPos = layerPos.ToString(),
                                LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                LayerType = "Annotation",
                                GroupLayerName = Common.EncloseStringInDoubleQuotes(layerContainer),
                                IsVisible = layer.IsVisible.ToString(),
                                LayerSource = annotationLayer.GetTable().GetPath().ToString(),
                                ClassName = annotationLayer.GetTable().GetName(),
                                IsSubtypeLayer = "FALSE",
                                GeometryType = annotationLayer.ShapeType.ToString(),
                                IsSelectable = annotationLayer.IsSelectable.ToString(),
                                IsEditable = annotationLayer.IsEditable.ToString(),
                                RefreshRate = cimAnnotationLayer.RefreshRate.ToString(),
                                DefinitionQuery = Common.EncloseStringInDoubleQuotes(annotationLayer.DefinitionFilter.DefinitionExpression),
                                MinScale = GetScaleValue(annotationLayer.MinScale),
                                MaxScale = GetScaleValue(annotationLayer.MaxScale)
                            };
                            layerInfoList.Add(rec);
                            increaseLayerPos = true;

                            foreach (var definitionFilter in annotationLayer.GetDefinitionFilters())
                            {
                                var definitionQueryInfo = new CsvDefinitionQueryInfo
                                {
                                    LayerPos = layerPos.ToString(),
                                    LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                    LayerType = "Annotation Layer",
                                    GroupLayerName = Common.EncloseStringInDoubleQuotes(layerContainer),
                                    DefinitionQueryName = definitionFilter.Name,
                                    DefinitionQuery = Common.EncloseStringInDoubleQuotes(definitionFilter.DefinitionExpression),
                                };
                                definitionQueryInfoList.Add(definitionQueryInfo);
                            }

                            CIMAnnotationLayer cimAnnotationLayerDef = layer.GetDefinition() as CIMAnnotationLayer;
                            if (cimAnnotationLayerDef.EnableDisplayFilters)
                            {
                                CIMDisplayFilter[] cimDisplayFilterChoices = cimAnnotationLayerDef.DisplayFilterChoices;
                                CIMDisplayFilter[] cimDisplayFilter = cimAnnotationLayerDef.DisplayFilters;
                                AddDisplayFiltersToList(rec, cimDisplayFilterChoices, cimDisplayFilter, ref layerInfoList);
                            }
                        }
                        else if (layer is AnnotationSubLayer annotationSubLayer)
                        {
                            CsvLayerInfo rec = new CsvLayerInfo()
                            {
                                LayerPos = layerPos.ToString(),
                                LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                LayerType = "Annotation Sub Layer",
                                GroupLayerName = Common.EncloseStringInDoubleQuotes(layerContainer),
                                IsVisible = layer.IsVisible.ToString(),
                                MinScale = GetScaleValue(layer.MinScale),
                                MaxScale = GetScaleValue(layer.MaxScale)
                            };

                            layerInfoList.Add(rec);
                            increaseLayerPos = true;
                        }
                        else if (layer is DimensionLayer dimensionLayer)
                        {
                            CIMDimensionLayer cimDimensionLayer = layer.GetDefinition() as CIMDimensionLayer;

                            CsvLayerInfo rec = new CsvLayerInfo()
                            {
                                LayerPos = layerPos.ToString(),
                                LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                LayerType = "Dimension",
                                GroupLayerName = Common.EncloseStringInDoubleQuotes(layerContainer),
                                IsVisible = layer.IsVisible.ToString(),
                                LayerSource = dimensionLayer.GetTable().GetPath().ToString(),
                                ClassName = dimensionLayer.GetTable().GetName(),
                                IsSubtypeLayer = "FALSE",
                                GeometryType = dimensionLayer.ShapeType.ToString(),
                                IsSelectable = dimensionLayer.IsSelectable.ToString(),
                                IsEditable = dimensionLayer.IsEditable.ToString(),
                                RefreshRate = cimDimensionLayer.RefreshRate.ToString(),
                                DefinitionQuery = Common.EncloseStringInDoubleQuotes(dimensionLayer.DefinitionFilter.DefinitionExpression),
                                MinScale = GetScaleValue(dimensionLayer.MinScale),
                                MaxScale = GetScaleValue(dimensionLayer.MaxScale)
                            };
                            layerInfoList.Add(rec);
                            increaseLayerPos = true;

                            foreach (var definitionFilter in dimensionLayer.GetDefinitionFilters())
                            {
                                var definitionQueryInfo = new CsvDefinitionQueryInfo
                                {
                                    LayerPos = layerPos.ToString(),
                                    LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                    LayerType = "Dimension Layer",
                                    GroupLayerName = Common.EncloseStringInDoubleQuotes(layerContainer),
                                    DefinitionQueryName = definitionFilter.Name,
                                    DefinitionQuery = Common.EncloseStringInDoubleQuotes(definitionFilter.DefinitionExpression),
                                };
                                definitionQueryInfoList.Add(definitionQueryInfo);
                            }

                            CIMDimensionLayer cimDimensionLayerDef = layer.GetDefinition() as CIMDimensionLayer;
                            if (cimDimensionLayerDef.EnableDisplayFilters)
                            {
                                CIMDisplayFilter[] cimDisplayFilterChoices = cimDimensionLayerDef.DisplayFilterChoices;
                                CIMDisplayFilter[] cimDisplayFilter = cimDimensionLayerDef.DisplayFilters;
                                AddDisplayFiltersToList(rec, cimDisplayFilterChoices, cimDisplayFilter, ref layerInfoList);
                            }

                        }
                        else if (layer is UtilityNetworkLayer utilityNetworkLayer)
                        {
                            CsvLayerInfo rec = new CsvLayerInfo()
                            {
                                LayerPos = layerPos.ToString(),
                                LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                GroupLayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                LayerType = "Utility Network Layer",
                                IsVisible = layer.IsVisible.ToString()
                            };
                            layerInfoList.Add(rec);

                            //Active Trace Configuration introduced in Utility Network version 5.
                            CIMUtilityNetworkLayer cimUtilityNetworkLayer = layer.GetDefinition() as CIMUtilityNetworkLayer;
                            CIMNetworkTraceConfiguration[] cimNetworkTraceConfigurations = cimUtilityNetworkLayer.ActiveTraceConfigurations;
                            if (cimNetworkTraceConfigurations != null)
                            {
                                for (int j = 0; j < cimNetworkTraceConfigurations.Length; j++)
                                {
                                    rec = new CsvLayerInfo()
                                    {
                                        LayerPos = layerPos.ToString(),
                                        LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                        GroupLayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                        ActiveTraceConfiguration = Common.EncloseStringInDoubleQuotes(cimNetworkTraceConfigurations[j].Name)
                                    };
                                    layerInfoList.Add(rec);
                                }
                            }
                            increaseLayerPos = true;
                        }
                        else if (layer is TiledServiceLayer tiledServiceLayer)
                        {
                            CsvLayerInfo rec = new CsvLayerInfo()
                            {
                                LayerPos = layerPos.ToString(),
                                LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                LayerType = "Tiled Service Layer",
                                IsVisible = layer.IsVisible.ToString(),
                                LayerSource = tiledServiceLayer.URL
                            };
                            layerInfoList.Add(rec);
                            increaseLayerPos = true;
                        }
                        else if (layer is VectorTileLayer vectorTileLayer)
                        {
                            CIMVectorTileDataConnection cimVectorTileDataConn = layer.GetDataConnection() as CIMVectorTileDataConnection;

                            CsvLayerInfo rec = new CsvLayerInfo()
                            {
                                LayerPos = layerPos.ToString(),
                                LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                LayerType = "Vector Tile Layer",
                                IsVisible = layer.IsVisible.ToString(),
                                LayerSource = cimVectorTileDataConn.URI
                            };
                            layerInfoList.Add(rec);
                            increaseLayerPos = true;
                        }
                        else if (layer is GraphicsLayer graphicsLayer)
                        {
                            CIMGraphicsLayer cimGraphicsLayer = layer.GetDefinition() as CIMGraphicsLayer;

                            CsvLayerInfo rec = new CsvLayerInfo()
                            {
                                LayerPos = layerPos.ToString(),
                                GroupLayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                LayerType = "Graphics Layer",
                                IsVisible = layer.IsVisible.ToString(),
                                IsSelectable = cimGraphicsLayer.Selectable.ToString(),
                                RefreshRate = cimGraphicsLayer.RefreshRate.ToString(),
                                MinScale = GetScaleValue(layer.MinScale),
                                MaxScale = GetScaleValue(layer.MaxScale)
                            };
                            layerInfoList.Add(rec);
                            increaseLayerPos = true;
                        }
                        else if (layer.MapLayerType == MapLayerType.BasemapBackground)
                        {
                            CsvLayerInfo rec = new CsvLayerInfo()
                            {
                                LayerPos = layerPos.ToString(),
                                LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                LayerType = "Basemap",
                                IsVisible = layer.IsVisible.ToString()
                            };
                            layerInfoList.Add(rec);
                            increaseLayerPos = true;
                        }
                        else
                        {
                            CsvLayerInfo rec = new CsvLayerInfo()
                            {
                                LayerPos = layerPos.ToString(),
                                LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                                LayerType = "Not Defined in this tool",
                                IsVisible = layer.IsVisible.ToString()
                            };
                            layerInfoList.Add(rec);
                            increaseLayerPos = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        CsvLayerInfo rec = new CsvLayerInfo()
                        {
                            LayerPos = layerPos.ToString(),
                            GroupLayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                            LayerName = "Extract Error",
                            IsVisible = layer.IsVisible.ToString(),
                            LayerSource = ex.Message
                        };
                        layerInfoList.Add(rec);
                        increaseLayerPos = true;
                    }

                    //increment counter by 1
                    if (increaseLayerPos)
                        layerPos += 1;

                    increaseLayerPos = false;
                }

                //Standalone Tables
                IReadOnlyList<StandaloneTable> standaloneTableList = MapView.Active.Map.StandaloneTables;
                foreach (StandaloneTable standaloneTable in standaloneTableList)
                {
                    CIMStandaloneTable cimStandaloneTable = standaloneTable.GetDefinition() as CIMStandaloneTable;
                    CIMExpressionInfo cimExpressionInfo = cimStandaloneTable.DisplayExpressionInfo;

                    //Primary Display Field
                    string displayField = cimStandaloneTable.DisplayField;
                    if (cimExpressionInfo != null)
                        displayField = cimExpressionInfo.Expression.Replace("\"", "'");  //double quotes messes up the delimeters in the CSV

                    CsvLayerInfo rec = new CsvLayerInfo()
                    {
                        LayerPos = layerPos.ToString(),
                        LayerName = Common.EncloseStringInDoubleQuotes(standaloneTable.Name),
                        LayerType = "Table",
                        LayerSource = standaloneTable.GetTable().GetPath().ToString(),
                        ClassName = standaloneTable.GetTable().GetName(),
                        DefinitionQuery = Common.EncloseStringInDoubleQuotes(standaloneTable.DefinitionFilter.DefinitionExpression),
                        DisplayField = Common.EncloseStringInDoubleQuotes(displayField),
                    };
                    layerInfoList.Add(rec);

                    foreach (var definitionFilter in standaloneTable.GetDefinitionFilters())
                    {
                        var definitionQueryInfo = new CsvDefinitionQueryInfo
                        {
                            LayerPos = layerPos.ToString(),
                            LayerName = Common.EncloseStringInDoubleQuotes(standaloneTable.Name),
                            LayerType = "Table",
                            GroupLayerName = Common.EncloseStringInDoubleQuotes(layerContainer),
                            DefinitionQueryName = definitionFilter.Name,
                            DefinitionQuery = Common.EncloseStringInDoubleQuotes(definitionFilter.DefinitionExpression),
                        };
                        definitionQueryInfoList.Add(definitionQueryInfo);
                    }

                    //Include Pop-up expressions if exist
                    if (cimStandaloneTable.PopupInfo != null)
                    {
                        if (cimStandaloneTable.PopupInfo.ExpressionInfos != null)
                        {
                            bool popupExprVisibility = false;
                            for (int i = 0; i < cimStandaloneTable.PopupInfo.ExpressionInfos.Length; i++)
                            {
                                //determine if expression is visible in popup
                                CIMMediaInfo[] cimMediaInfos = cimStandaloneTable.PopupInfo.MediaInfos;
                                for (int j = 0; j < cimMediaInfos.Length; j++)
                                {
                                    if (cimMediaInfos[j] is CIMTableMediaInfo cimTableMediaInfo)
                                    {
                                        string[] fields = cimTableMediaInfo.Fields;
                                        for (int k = 0; k < fields.Length; k++)
                                        {
                                            if (fields[k] == "expression/" + cimStandaloneTable.PopupInfo.ExpressionInfos[i].Name)
                                            {
                                                popupExprVisibility = true;
                                                break;
                                            }
                                        }
                                    }

                                    //Break out of 2nd loop (j) if already found the expression
                                    if (popupExprVisibility)
                                        break;
                                }

                                //Write popup info
                                var popupRec = new CsvPopupInfo()
                                {
                                    LayerPos = layerPos.ToString(),
                                    LayerName = Common.EncloseStringInDoubleQuotes(standaloneTable.Name),
                                    LayerType = "Table",
                                    PopupExpresssionName = cimStandaloneTable.PopupInfo.ExpressionInfos[i].Name,
                                    PopupExpresssionTitle = Common.EncloseStringInDoubleQuotes(cimStandaloneTable.PopupInfo.ExpressionInfos[i].Title.Replace("\"", "'")),
                                    PopupExpresssionVisible = popupExprVisibility.ToString(),
                                    PopupExpressionArcade = Common.EncloseStringInDoubleQuotes(cimStandaloneTable.PopupInfo.ExpressionInfos[i].Expression.Replace("\"", "'"))
                                };
                                popupInfoList.Add(popupRec);
                            }
                        }
                    }

                    layerPos += 1;
                }

                var headerInfo = new string[] {
                    DateTime.Now + "," + "Layer Info",
                    "",
                    "Project," + Project.Current.Path,
                    "Map," + MapView.Active.Map.Name,
                    "Layer Count," + MapView.Active.Map.GetLayersAsFlattenedList().OfType<Layer>().Count(),
                    "Table Count," + MapView.Active.Map.StandaloneTables.Count,
                    "",
                };


                if (layerInfoList.Count > 0)
                    using (var sw = new StreamWriter(outputLayerFileName))
                    {
                        //Header information
                        foreach (var line in headerInfo)
                            sw.WriteLine(line);

                        //Get all properties defined in the class.  This will be used to generate the CSV file
                        //Write column headers based on properties in the class
                        var csvProperties = Common.GetPropertiesOfClass(new CsvLayerInfo());
                        sw.WriteLine(Common.ExtractClassPropertyNamesToString(csvProperties));
                        foreach (CsvLayerInfo row in layerInfoList)
                            sw.WriteLine(Common.ExtractClassValuesToString(row, csvProperties));

                        sw.Flush();
                        sw.Close();
                    }

                if(labelInfoList.Count > 0)
                    using(var sw = new StreamWriter(outputLabelFileName))
                    {
                        //Header information
                        foreach (var line in headerInfo)
                            sw.WriteLine(line);

                        //Get all properties defined in the class.  This will be used to generate the CSV file
                        //Write column headers based on properties in the class
                        var csvProperties = Common.GetPropertiesOfClass(new CsvLabelInfo());
                        sw.WriteLine(Common.ExtractClassPropertyNamesToString(csvProperties));
                        foreach (var row in labelInfoList)
                            sw.WriteLine(Common.ExtractClassValuesToString(row, csvProperties));

                        sw.Flush();
                        sw.Close();

                    }
                if (popupInfoList.Count > 0)
                    using (var sw = new StreamWriter(outputPopupFileName))
                    {
                        //Header information
                        foreach (var line in headerInfo)
                            sw.WriteLine(line);

                        //Get all properties defined in the class.  This will be used to generate the CSV file
                        //Write column headers based on properties in the class
                        var csvProperties = Common.GetPropertiesOfClass(new CsvPopupInfo());
                        sw.WriteLine(Common.ExtractClassPropertyNamesToString(csvProperties));
                        foreach (var row in popupInfoList)
                            sw.WriteLine(Common.ExtractClassValuesToString(row, csvProperties));

                        sw.Flush();
                        sw.Close();

                    }
                if (definitionQueryInfoList.Count > 0)
                    using (var sw = new StreamWriter(outputDefQeryFileName))
                    {
                        //Header information
                        foreach (var line in headerInfo)
                            sw.WriteLine(line);

                        //Get all properties defined in the class.  This will be used to generate the CSV file
                        //Write column headers based on properties in the class
                        var csvProperties = Common.GetPropertiesOfClass(new CsvDefinitionQueryInfo());
                        sw.WriteLine(Common.ExtractClassPropertyNamesToString(csvProperties));
                        foreach (var row in definitionQueryInfoList)
                            sw.WriteLine(Common.ExtractClassValuesToString(row, csvProperties));

                        sw.Flush();
                        sw.Close();
                    }
            });
        }

        private static void AddDisplayFiltersToList(CsvLayerInfo parentRec, CIMDisplayFilter[] cimDisplayFilterChoices, CIMDisplayFilter[] cimDisplayFilter, ref List<CsvLayerInfo> layerAttributeList)
        {
            //TODO::Separate list
            //In Pro, there are 2 choices to set the Active Display Filters
            //option 1:  Manually 
            if (cimDisplayFilterChoices != null)
            {
                for (int j = 0; j < cimDisplayFilterChoices.Length; j++)
                {
                    CsvLayerInfo rec = new CsvLayerInfo()
                    {
                        LayerPos = parentRec.LayerPos,
                        GroupLayerName = parentRec.GroupLayerName,
                        DisplayFilterName = Common.EncloseStringInDoubleQuotes(cimDisplayFilterChoices[j].Name),
                        DisplayFilterExpresssion = Common.EncloseStringInDoubleQuotes(cimDisplayFilterChoices[j].WhereClause),
                    };
                    layerAttributeList.Add(rec);
                }
            }

            //option 2:  By Scale
            if (cimDisplayFilter != null)
            {
                for (int k = 0; k < cimDisplayFilter.Length; k++)
                {
                    CsvLayerInfo rec = new CsvLayerInfo()
                    {
                        LayerPos = parentRec.LayerPos,
                        GroupLayerName = parentRec.GroupLayerName,
                        DisplayFilterName = Common.EncloseStringInDoubleQuotes(cimDisplayFilter[k].Name),
                        MinScale = GetScaleValue(cimDisplayFilter[k].MinScale),
                        MaxScale = GetScaleValue(cimDisplayFilter[k].MaxScale)
                    };
                    layerAttributeList.Add(rec);
                }
            }
        }

        private static void DetermineSymbology(CIMFeatureLayer cimFeatureLayerDef, out string primarySymbology, out string field1, out string field2, out string field3)
        {
            primarySymbology = string.Empty;
            field1 = string.Empty;
            field2 = string.Empty;
            field3 = string.Empty;

            //Symbology
            if (cimFeatureLayerDef.Renderer is CIMSimpleRenderer)
                primarySymbology = "Single Symbol";
            else if (cimFeatureLayerDef.Renderer is CIMUniqueValueRenderer uniqueRenderer)
            {
                primarySymbology = "Unique Values";

                switch (uniqueRenderer.Fields.Length)
                {
                    case 1:
                        field1 = uniqueRenderer.Fields[0];
                        break;
                    case 2:
                        field1 = uniqueRenderer.Fields[0];
                        field2 = uniqueRenderer.Fields[1];
                        break;
                    case 3:
                        field1 = uniqueRenderer.Fields[0];
                        field2 = uniqueRenderer.Fields[1];
                        field3 = uniqueRenderer.Fields[2];
                        break;
                }
            }
            else if (cimFeatureLayerDef.Renderer is CIMChartRenderer)
                primarySymbology = "Charts";
            else if (cimFeatureLayerDef.Renderer is CIMClassBreaksRendererBase classBreaksRenderer)
                primarySymbology = classBreaksRenderer.ClassBreakType.ToString();
            else if (cimFeatureLayerDef.Renderer is CIMDictionaryRenderer)
                primarySymbology = "Dictionary";
            else if (cimFeatureLayerDef.Renderer is CIMDotDensityRenderer)
                primarySymbology = "Dot Density";
            else if (cimFeatureLayerDef.Renderer is CIMHeatMapRenderer)
                primarySymbology = "Heat Map";
            else if (cimFeatureLayerDef.Renderer is CIMProportionalRenderer)
                primarySymbology = "Proportional Symbols";
            else if (cimFeatureLayerDef.Renderer is CIMRepresentationRenderer)
                primarySymbology = "Representation";
        }

        private static string GetScaleValue(double scale)
        {
            if (scale == 0)
                return "<None>";  // In Pro, when there is no scale set, the value is null.  Thru the SDK, it was showing 0.
            else
                return scale.ToString();
        }

        private class CsvLayerInfo
        {
            public string LayerPos { get; set; }
            public string LayerType { get; set; }
            public string GroupLayerName { get; set; }
            public string LayerName { get; set; }
            public string IsVisible { get; set; }
            public string LayerSource { get; set; }
            public string ClassName { get; set; }
            public string IsSubtypeLayer { get; set; }
            public string SubtypeValue { get; set; }
            public string GeometryType { get; set; }
            public string IsSnappable { get; set; }
            public string IsSelectable { get; set; }
            public string IsEditable { get; set; }
            public string RefreshRate { get; set; }
            public string ActiveTraceConfiguration { get; set; }
            public string DefinitionQueryName { get; set; }
            public string DefinitionQuery { get; set; }
            public string DisplayFilterName { get; set; }
            public string DisplayFilterExpresssion { get; set; }
            public string MinScale { get; set; }
            public string MaxScale { get; set; }
            public string ShowMapTips { get; set; }
            public string PrimarySymbology { get; set; }
            public string SymbologyField1 { get; set; }
            public string SymbologyField2 { get; set; }
            public string SymbologyField3 { get; set; }
            public string EditTemplateCount { get; set; }
            public string DisplayField { get; set; }
            public string IsLabelVisible { get; set; }
            public string LabelExpression { get; set; }
            public string LabelMinScale { get; set; }
            public string LabelMaxScale { get; set; }
            public string PopupExpresssionName { get; set; }
            public string PopupExpresssionTitle { get; set; }
            public string PopupExpresssionVisible { get; set; }
            public string PopupExpressionArcade { get; set; }
        }

        private class CsvDefinitionQueryInfo
        {
            public string LayerPos { get; set; }
            public string LayerType { get; set; }
            public string GroupLayerName { get; set; }
            public string LayerName { get; set; }
            public string DefinitionQueryName { get; set; }
            public string DefinitionQuery { get; set; }
        }

        private class CsvLabelInfo
        {
            public string LayerPos { get; set; }
            public string LayerType { get; set; }
            public string GroupLayerName { get; set; }
            public string LayerName { get; set; }
            public string LabelClassName { get; set; }
            public string IsLabelVisible { get; set; }
            public string LabelExpression { get; set; }
            public string LabelMinScale { get; set; }
            public string LabelMaxScale { get; set; }
        }
        private class CsvPopupInfo
        {
            public string LayerPos { get; set; }
            public string LayerType { get; set; }
            public string GroupLayerName { get; set; }
            public string LayerName { get; set; }
            public string PopupExpresssionName { get; set; }
            public string PopupExpresssionTitle { get; set; }
            public string PopupExpresssionVisible { get; set; }
            public string PopupExpressionArcade { get; set; }
        }
    }
}
