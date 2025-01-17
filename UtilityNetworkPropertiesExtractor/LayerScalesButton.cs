﻿/*
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
    internal class LayerScalesButton : Button
    {
        private static string _fileName = string.Empty;

        protected async override void OnClick()
        {
            try
            {
                await ExtractLayerScalesAsync();
                MessageBox.Show("Directory: " + Common.ExtractFilePath + Environment.NewLine + "File Name: " + _fileName, "CSV file has been generated");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Extract Layer Info");
            }
        }

        public static Task ExtractLayerScalesAsync()
        {
            return QueuedTask.Run(() =>
            {
                Common.CreateOutputDirectory();

                string dateFormatted = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                _fileName = string.Format("{0}_{1}_LayerScales.csv", dateFormatted, Common.GetProProjectName());
                string outputFile = Path.Combine(Common.ExtractFilePath, _fileName);

                using (StreamWriter sw = new StreamWriter(outputFile))
                {
                    //Header information
                    sw.WriteLine(DateTime.Now + "," + "Layer Scales");
                    sw.WriteLine();
                    sw.WriteLine("Project," + Project.Current.Path);
                    sw.WriteLine("Map," + MapView.Active.Map.Name);
                    sw.WriteLine("Layer Count," + MapView.Active.Map.GetLayersAsFlattenedList().OfType<Layer>().Count());
                    sw.WriteLine();

                    //Since you can't name a C#.NET property with numeric values, I spelled them out.
                    //Then populate a record with the numeric equivalent and use that to write the header
                    CSVLayout header = new CSVLayout()
                    {
                        LayerPos = "Pos",
                        LayerType = "LayerType",
                        GroupLayerName = "GroupLayerName",
                        LayerName = "LayerName",
                        ScaleRange = "ScaleRange",
                        Zero = "0",
                        FiveHundred = "500",
                        TwelveHundred = "1200",
                        TwentyFiveHundred = "2500",
                        FiveThousand = "5000",
                        TenThousand = "10000",
                        FiftyThousand = "50000",
                        OneHundredThousand = "100000",
                        TwoHundredThousand = "200000",
                        OneMillion = "1000000",
                        TenMillion = "10000000"
                    };

                    PropertyInfo[] properties = Common.GetPropertiesOfClass(header);
                    string output = Common.ExtractClassValuesToString(header, properties);
                    sw.WriteLine(output);

                    List<CSVLayout> CSVLayoutList = new List<CSVLayout>();
                    
                    int layerPos = 1;
                    string groupLayerName = string.Empty;
                    string prevGroupLayerName = string.Empty;
                    string layerContainer = string.Empty;
                    string layerType = string.Empty;

                    List<Layer> layerList = MapView.Active.Map.GetLayersAsFlattenedList().OfType<Layer>().ToList();
                    foreach (Layer layer in layerList)
                    {
                        //Determine if in a group layer
                        layerContainer = layer.Parent.ToString();
                        if (layerContainer != MapView.Active.Map.Name) // Group layer
                        {
                            if (layerContainer != prevGroupLayerName)
                                prevGroupLayerName = layerContainer;
                        }
                        else
                            layerContainer = string.Empty;

                        layerType = GetLayerType(layer);
                        switch (layerType) {
                            case "Feature Layer":
                            case "Annotation":
                            case "Annotation Sub Layer":
                            case "Dimension":
                                groupLayerName = Common.EncloseStringInDoubleQuotes(layerContainer);
                                break;
                            default:
                                groupLayerName = Common.EncloseStringInDoubleQuotes(layer.Name);
                                break;
                        }

                        //Set values for the layer
                        CSVLayout scaleRec = new CSVLayout()
                        {
                            LayerPos = layerPos.ToString(),
                            LayerType = layerType,
                            GroupLayerName = groupLayerName,
                            LayerName = Common.EncloseStringInDoubleQuotes(layer.Name),
                            ScaleRange = Common.EncloseStringInDoubleQuotes(layer.MaxScale + " -- " + layer.MinScale),
                            Zero = IsLayerRenderedAtThisScale(header.Zero, layer).ToString(),
                            FiveHundred = IsLayerRenderedAtThisScale(header.FiveHundred, layer).ToString(),
                            TwelveHundred = IsLayerRenderedAtThisScale(header.TwelveHundred, layer).ToString(),
                            TwentyFiveHundred = IsLayerRenderedAtThisScale(header.TwentyFiveHundred, layer).ToString(),
                            FiveThousand = IsLayerRenderedAtThisScale(header.FiveThousand, layer).ToString(),
                            TenThousand = IsLayerRenderedAtThisScale(header.TenThousand, layer).ToString(),
                            FiftyThousand = IsLayerRenderedAtThisScale(header.FiftyThousand, layer).ToString(),
                            OneHundredThousand = IsLayerRenderedAtThisScale(header.OneHundredThousand, layer).ToString(),
                            TwoHundredThousand = IsLayerRenderedAtThisScale(header.TwoHundredThousand, layer).ToString(),
                            OneMillion = IsLayerRenderedAtThisScale(header.OneMillion, layer).ToString(),
                            TenMillion = IsLayerRenderedAtThisScale(header.TenMillion, layer).ToString()
                        };

                        CSVLayoutList.Add(scaleRec);
                        layerPos += 1;
                    }

                    foreach (CSVLayout row in CSVLayoutList)
                    {
                        output = Common.ExtractClassValuesToString(row, properties);
                        sw.WriteLine(output);
                    }

                    sw.Flush();
                    sw.Close();
                }
            });
        }

        private static bool IsLayerRenderedAtThisScale(string scaleText, Layer layer)
        {
            bool retVal = false;

            if (layer.MinScale == 0 && layer.MaxScale == 0)  //Min and Max scale weren't defined.  Layer will renderer at any scale
                retVal = true;
            else
            {
                double scale = Convert.ToDouble(scaleText);
                if (layer.MinScale >= scale && layer.MaxScale <= scale)
                    retVal = true;
            }
            return retVal;
        }

        private static string GetLayerType(Layer layer)
        {
            string retVal;

            if (layer is FeatureLayer)
                retVal = "Feature Layer";
            else if (layer is GroupLayer)
                retVal = "Group Layer";
            else if (layer is SubtypeGroupLayer)
                retVal = "Subtype Group Layer";
            else if (layer is AnnotationLayer)
                retVal = "Annotation";
            else if (layer is AnnotationSubLayer)
                retVal = "Annotation Sub Layer";
            else if (layer is DimensionLayer)
                retVal = "Dimension";
            else if (layer is UtilityNetworkLayer)
                retVal = "Utility Network Layer";
            else if (layer is TiledServiceLayer)
                retVal = "Tiled Service Layer";
            else if (layer is VectorTileLayer)
                retVal = "Vector Tile Layer";
            else if (layer is GraphicsLayer)
                retVal = "Graphics Layer";
            else
                retVal = "Layer";

            return retVal;
        }

        private class CSVLayout
        {
            public string LayerPos { get; set; }
            public string LayerType { get; set; }
            public string GroupLayerName { get; set; }
            public string LayerName { get; set; }
            public string ScaleRange { get; set; }
            public string Zero { get; set; }
            public string FiveHundred { get; set; }
            public string TwelveHundred { get; set; }
            public string TwentyFiveHundred { get; set; }
            public string FiveThousand { get; set; }
            public string TenThousand { get; set; }
            public string FiftyThousand { get; set; }
            public string OneHundredThousand { get; set; }
            public string TwoHundredThousand { get; set; }
            public string OneMillion { get; set; }
            public string TenMillion { get; set; }
        }
    }
}