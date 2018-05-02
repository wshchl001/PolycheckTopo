﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geoprocessor;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.Display;

namespace PolycheckTopo.checkAreaPointTopo
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }
        string AltDirectorySeparator = System.IO.Path.AltDirectorySeparatorChar.ToString();
        ILayer selectedAreaLayer;
        ILayer selectedPointLayer;
        bool areacheck, pointcheck = false;

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedAreaLayer = layersList[comboBox1.SelectedIndex];
            bool flag1 = (selectedAreaLayer as IFeatureLayer).FeatureClass.ShapeType == ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon;
            areacheck = flag1;
            if (pointcheck && areacheck)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedPointLayer = layersList[comboBox2.SelectedIndex];
            bool flag2 = (selectedPointLayer as IFeatureLayer).FeatureClass.ShapeType == ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint;
            pointcheck = flag2;
            if (areacheck && pointcheck)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool b = ((selectedAreaLayer as IDataset).Workspace as IWorkspaceEdit).IsBeingEdited();
            bool b1 = ((selectedPointLayer as IDataset).Workspace as IWorkspaceEdit).IsBeingEdited();
            if (b && b1)
            {
                MessageBox.Show("请关闭编辑");
                return;
            }

            this.BeginInvoke(
                new Action(() =>
                {
                    Invoke(new Action(() =>
                    {
                        this.button1.Enabled = false;
                        this.ControlBox = false;
                        this.comboBox1.Enabled = false;
                        this.comboBox2.Enabled = false;
                        isworkingflag = true;
                    }));
                    start();
                    Invoke(new Action(() =>
                    {
                        this.ControlBox = true;
                        this.button1.Enabled = true;
                        this.comboBox1.Enabled = true;
                        this.comboBox2.Enabled = true;
                        isworkingflag = false;
                    }));
                }));
            MessageBox.Show("完成,请查看拓扑错误与界址点中NEAR_DIST字段中的值(为离该点最近点的距离),多余字段不必删除,不影响数据结构,或者手动移除JZD所有连接。");

        }
        bool isworkingflag;
        protected override void OnClosing(CancelEventArgs e)
        {
            if (isworkingflag == true)
                e.Cancel = true;
            base.OnClosing(e);
        }
        List<ILayer> layersList;
        void start()
        {
            string datasetname;
            string workspacepathnameofDK = (selectedAreaLayer as IDataset).Workspace.PathName;
            string workspacepathnameofJZD = (selectedPointLayer as IDataset).Workspace.PathName;
            string DKname = (selectedAreaLayer as IDataset).BrowseName;
            string DKpath = workspacepathnameofDK + AltDirectorySeparator + DKname;
            string JZDname = (selectedPointLayer as IDataset).BrowseName;
            string JZDpath = workspacepathnameofJZD + AltDirectorySeparator + JZDname;
            string DKnewname;
            string JZDnewname;


            //创建要素数据集
            Geoprocessor gp = new Geoprocessor();
            gp.OverwriteOutput = true;
            gp.AddOutputsToMap = false;
            ESRI.ArcGIS.DataManagementTools.CreateFeatureDataset cfd = new ESRI.ArcGIS.DataManagementTools.CreateFeatureDataset();
            cfd.out_dataset_path = workspacepathnameofDK;
            cfd.spatial_reference = DKpath;
            cfd.out_name = datasetname = "T" + Guid.NewGuid().ToString("N");
            gp.Execute(cfd, null);
            Invoke(new Action(() => { progressBar1.Value = 10; }));

            //将DK复制到要素数据集中
            ESRI.ArcGIS.DataManagementTools.Copy copy = new ESRI.ArcGIS.DataManagementTools.Copy();
            copy.in_data = DKpath;
            DKnewname = "T" + Guid.NewGuid().ToString("N");
            copy.out_data = workspacepathnameofDK + AltDirectorySeparator + datasetname + AltDirectorySeparator + DKnewname;
            gp.Execute(copy, null);
            Invoke(new Action(() => { progressBar1.Value = 20; }));


            //将JZD复制到要素数据集中
            ESRI.ArcGIS.DataManagementTools.Copy copy1 = new ESRI.ArcGIS.DataManagementTools.Copy();
            copy1.in_data = JZDpath;
            JZDnewname = "T" + Guid.NewGuid().ToString("N");
            copy1.out_data = workspacepathnameofDK + AltDirectorySeparator + datasetname + AltDirectorySeparator + JZDnewname;
            gp.Execute(copy1, null);
            Invoke(new Action(() => { progressBar1.Value = 30; }));


            //添加拓扑
            ESRI.ArcGIS.DataManagementTools.CreateTopology ctp = new ESRI.ArcGIS.DataManagementTools.CreateTopology();
            ctp.in_dataset = workspacepathnameofDK + AltDirectorySeparator + datasetname;
            ctp.in_cluster_tolerance = 0.001;
            ctp.out_name = datasetname + "topo";
            gp.Execute(ctp, null);
            Invoke(new Action(() => { progressBar1.Value = 40; }));


            //向拓扑中添加要素类
            ESRI.ArcGIS.DataManagementTools.AddFeatureClassToTopology afct = new ESRI.ArcGIS.DataManagementTools.AddFeatureClassToTopology();
            afct.in_featureclass = copy.out_data;
            afct.in_topology = workspacepathnameofDK + AltDirectorySeparator + datasetname + AltDirectorySeparator + ctp.out_name;
            gp.Execute(afct, null);
            ESRI.ArcGIS.DataManagementTools.AddFeatureClassToTopology afct1 = new ESRI.ArcGIS.DataManagementTools.AddFeatureClassToTopology();
            afct1.in_featureclass = copy1.out_data;
            afct1.in_topology = workspacepathnameofDK + AltDirectorySeparator + datasetname + AltDirectorySeparator + ctp.out_name;
            gp.Execute(afct1, null);


            //添加拓扑规则
            ESRI.ArcGIS.DataManagementTools.AddRuleToTopology art = new ESRI.ArcGIS.DataManagementTools.AddRuleToTopology();
            art.in_featureclass = afct1.in_featureclass;
            art.in_featureclass2 = afct.in_featureclass;
            art.in_topology = afct.in_topology;
            art.rule_type = "Must Be Covered By Boundary Of (Point-Area)";
            gp.Execute(art, null);
            Invoke(new Action(() => { progressBar1.Value = 50; }));

            //验证拓扑
            ESRI.ArcGIS.DataManagementTools.ValidateTopology vt = new ESRI.ArcGIS.DataManagementTools.ValidateTopology();
            vt.in_topology = art.in_topology;
            gp.Execute(vt, null);
            Invoke(new Action(() => { progressBar1.Value = 60; }));


            //导出拓扑错误
            ESRI.ArcGIS.DataManagementTools.ExportTopologyErrors ete = new ESRI.ArcGIS.DataManagementTools.ExportTopologyErrors();
            ete.in_topology = vt.in_topology;
            ete.out_path = "in_memory" + AltDirectorySeparator;//输出结果保存在内存中
            ete.out_basename = "TopologicalError" + "T" + Guid.NewGuid().ToString("N");
            gp.AddOutputsToMap = true;
            gp.Execute(ete, null);
            Invoke(new Action(() => { progressBar1.Value = 70; }));


            
            IEnumLayer pls = ArcMap.Document.FocusMap.Layers;
            ILayer currentlayer;
            while ((currentlayer = pls.Next()) != null)
            {
                if (currentlayer.Name == ete.out_basename + "_point")
                {
                    //渲染线错误图层
                    IFeatureLayer featureLayer = currentlayer as IFeatureLayer;
                    IGeoFeatureLayer geoFeatureLayer = (IGeoFeatureLayer)featureLayer;
                    ISimpleRenderer simpleRenderer;
                    //ISimpleLineSymbol simpleLineSymbol = new SimpleLineSymbolClass();
                    ISimpleMarkerSymbol simpleLineSymbol = new SimpleMarkerSymbol();
                    simpleLineSymbol.Style = esriSimpleMarkerStyle.esriSMSSquare;
                    simpleLineSymbol.Size = 2.5;
                    IRgbColor rgbColor = new RgbColorClass();
                    rgbColor.Red = 255;
                    rgbColor.Green = 0;
                    rgbColor.Blue = 0;
                    simpleLineSymbol.Color = rgbColor;
                    simpleRenderer = new SimpleRendererClass();
                    simpleRenderer.Symbol = (ISymbol)simpleLineSymbol;
                    geoFeatureLayer.Renderer = (IFeatureRenderer)simpleRenderer;
                    ArcMap.Document.ActiveView.ContentsChanged();
                    ArcMap.Document.ActiveView.Refresh();

                }
                if (currentlayer.Name == ete.out_basename + "_poly")
                {
                    ArcMap.Document.FocusMap.DeleteLayer(currentlayer);
                }
                if (currentlayer.Name == ete.out_basename + "_line")
                {
                    ArcMap.Document.FocusMap.DeleteLayer(currentlayer);
                }
                if (currentlayer.Name == (selectedPointLayer as IDataset).Name)
                {
                    ArcMap.Document.FocusMap.DeleteLayer(currentlayer);
                }
                if (currentlayer.Name == (selectedAreaLayer as IDataset).Name)
                {
                    ArcMap.Document.FocusMap.DeleteLayer(currentlayer);
                }
            }

            //删除地块
            ESRI.ArcGIS.DataManagementTools.Delete d = new ESRI.ArcGIS.DataManagementTools.Delete();
            gp.AddOutputsToMap = false;
            d.in_data = DKpath;
            gp.Execute(d, null);
            //删除JZD
            ESRI.ArcGIS.DataManagementTools.Delete d1 = new ESRI.ArcGIS.DataManagementTools.Delete();
            gp.AddOutputsToMap = false;
            d1.in_data = JZDpath;
            gp.Execute(d1, null);
            Invoke(new Action(() => { progressBar1.Value = 80; }));

            //还原地块
            ESRI.ArcGIS.DataManagementTools.Copy c = new ESRI.ArcGIS.DataManagementTools.Copy();
            gp.AddOutputsToMap = false;
            c.in_data = copy.out_data;
            c.out_data = DKpath;
            gp.Execute(c, null);
            //还原JZX
            ESRI.ArcGIS.DataManagementTools.Copy c1 = new ESRI.ArcGIS.DataManagementTools.Copy();
            gp.AddOutputsToMap = false;
            c1.in_data = copy1.out_data;
            c1.out_data = JZDpath;
            gp.Execute(c1, null);
            Invoke(new Action(() => { progressBar1.Value = 90; }));

            IWorkspaceFactory w = new AccessWorkspaceFactoryClass();
            IWorkspace ws = w.OpenFromFile(workspacepathnameofDK, 0);
            IFeatureClass pFC = (ws as IFeatureWorkspace).OpenFeatureClass(DKname);
            IFeatureLayer pFLayer = new FeatureLayerClass();
            pFLayer.FeatureClass = pFC;
            pFLayer.Name = pFC.AliasName;
            ILayer pLayer = pFLayer as ILayer;
            IMap pMap = ArcMap.Document.FocusMap;
            ArcMap.Document.AddLayer(pLayer);


            w = new AccessWorkspaceFactoryClass();
            ws = w.OpenFromFile(workspacepathnameofJZD, 0);
            pFC = (ws as IFeatureWorkspace).OpenFeatureClass(JZDname);
            pFLayer = new FeatureLayerClass();
            pFLayer.FeatureClass = pFC;
            pFLayer.Name = pFC.AliasName;
            pLayer = pFLayer as ILayer;
            pMap = ArcMap.Document.FocusMap;
            ArcMap.Document.AddLayer(pLayer);

            //删除拓扑
            ESRI.ArcGIS.DataManagementTools.Delete dl = new ESRI.ArcGIS.DataManagementTools.Delete();
            gp.AddOutputsToMap = false;
            dl.in_data = ctp.in_dataset;
            gp.Execute(dl, null);


            ESRI.ArcGIS.AnalysisTools.GenerateNearTable gnt = new ESRI.ArcGIS.AnalysisTools.GenerateNearTable();
            gnt.in_features = JZDpath;
            gnt.near_features = JZDpath;
            gnt.out_table = "in_memory" + AltDirectorySeparator + "界址点临近分析T" + Guid.NewGuid().ToString("N");    
            gp.Execute(gnt, null);

            ESRI.ArcGIS.DataManagementTools.AddJoin aj = new ESRI.ArcGIS.DataManagementTools.AddJoin();
            aj.in_layer_or_view = pLayer;
            aj.in_field = "OBJECTID";
            aj.join_table = gnt.out_table;
            aj.join_field = "IN_FID";

            gp.Execute(aj, null);



            ArcMap.Document.ActiveView.ContentsChanged();
            ArcMap.Document.ActiveView.Refresh();

            Invoke(new Action(() => { progressBar1.Value = 100; }));



        }
        private void Form3_Load(object sender, EventArgs e)
        {
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            layersList = new List<ILayer>();
            IEnumLayer pls = ArcMap.Document.FocusMap.Layers;
            ILayer pl;
            while ((pl = pls.Next()) != null)
            {
                layersList.Add(pl);
            }
            layersList.ForEach(new Action<ILayer>((player) =>
            {
                comboBox1.Items.Add(player.Name);
                comboBox2.Items.Add(player.Name);
            }));
            button1.Enabled = false;

        }


    }
}
