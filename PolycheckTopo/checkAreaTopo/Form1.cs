using System;
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

namespace PolycheckTopo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        bool isworkingflag;
        string AltDirectorySeparator = System.IO.Path.AltDirectorySeparatorChar.ToString();
        string datasetname;
        string DKnewname;
        void start()
        {

            //IGeoProcessorResult pgr;
            //创建要素数据集
            Geoprocessor gp = new Geoprocessor();
            gp.OverwriteOutput = true;
            gp.AddOutputsToMap = false;
            ESRI.ArcGIS.DataManagementTools.CreateFeatureDataset cfd = new ESRI.ArcGIS.DataManagementTools.CreateFeatureDataset();
            string workspacepathname = (selectedLayer as IDataset).Workspace.PathName;
            cfd.out_dataset_path = workspacepathname;
            string DKname = (selectedLayer as IDataset).BrowseName;
            string DKpath = workspacepathname + AltDirectorySeparator + DKname;
            cfd.spatial_reference = DKpath;
            cfd.out_name = datasetname = "T" + Guid.NewGuid().ToString("N");
            gp.Execute(cfd, null);
            Invoke(new Action(() => { progressBar1.Value = 10; }));


            //将DK复制到要素数据集中
            ESRI.ArcGIS.DataManagementTools.Copy copy = new ESRI.ArcGIS.DataManagementTools.Copy();
            copy.in_data = DKpath;
            DKnewname = "T" + Guid.NewGuid().ToString("N");
            copy.out_data = workspacepathname + AltDirectorySeparator + datasetname + AltDirectorySeparator + DKnewname;
            gp.Execute(copy, null);
            Invoke(new Action(() => { progressBar1.Value = 20; }));

            //添加拓扑
            ESRI.ArcGIS.DataManagementTools.CreateTopology ctp = new ESRI.ArcGIS.DataManagementTools.CreateTopology();
            ctp.in_dataset = workspacepathname + AltDirectorySeparator + datasetname;
            ctp.in_cluster_tolerance = 0.001;
            ctp.out_name = datasetname + "topo";
            gp.Execute(ctp, null);

            Invoke(new Action(() => { progressBar1.Value = 30; }));

            //向拓扑中添加要素类
            ESRI.ArcGIS.DataManagementTools.AddFeatureClassToTopology afct = new ESRI.ArcGIS.DataManagementTools.AddFeatureClassToTopology();
            afct.in_featureclass = copy.out_data;
            afct.in_topology = workspacepathname + AltDirectorySeparator + datasetname + AltDirectorySeparator + ctp.out_name;
            gp.Execute(afct, null);

            Invoke(new Action(() => { progressBar1.Value = 40; }));

            //添加拓扑规则
            ESRI.ArcGIS.DataManagementTools.AddRuleToTopology art = new ESRI.ArcGIS.DataManagementTools.AddRuleToTopology();
            art.in_featureclass = afct.in_featureclass;
            art.in_topology = afct.in_topology;
            art.rule_type = "Must Not Have Gaps (Area)";
            gp.Execute(art, null);


            ESRI.ArcGIS.DataManagementTools.AddRuleToTopology art1 = new ESRI.ArcGIS.DataManagementTools.AddRuleToTopology();
            art1.in_featureclass = afct.in_featureclass;
            art1.in_topology = afct.in_topology;
            art1.rule_type = "Must Not Overlap (Area)";
            gp.Execute(art1, null);

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
            Invoke(new Action(() => { progressBar1.Value = 70; }));

            ete.out_basename = "TopologicalError" + "T" + Guid.NewGuid().ToString("N");
            gp.AddOutputsToMap = true;
            gp.Execute(ete, null);


            IEnumLayer pls = ArcMap.Document.FocusMap.Layers;
            ILayer currentlayer;
            while ((currentlayer = pls.Next()) != null)
            {
                if (currentlayer.Name == ete.out_basename + "_point")
                {
                    ArcMap.Document.FocusMap.DeleteLayer(currentlayer);
                }
                if (currentlayer.Name == ete.out_basename + "_poly")
                {
                    ArcMap.Document.FocusMap.DeleteLayer(currentlayer);
                }
                if (currentlayer.Name == ete.out_basename + "_line")
                {
                    //渲染线错误图层
                    IFeatureLayer featureLayer = currentlayer as IFeatureLayer;
                    IGeoFeatureLayer geoFeatureLayer = (IGeoFeatureLayer)featureLayer;
                    ISimpleRenderer simpleRenderer;
                    ISimpleLineSymbol simpleLineSymbol = new SimpleLineSymbolClass();
                    simpleLineSymbol.Style = esriSimpleLineStyle.esriSLSSolid;
                    simpleLineSymbol.Width = 2.5;
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
                if (currentlayer.Name == (selectedLayer as IDataset).Name)
                {
                    ArcMap.Document.FocusMap.DeleteLayer(currentlayer);
                }

            }
            //还原现场

            //删除地块
            ESRI.ArcGIS.DataManagementTools.Delete d = new ESRI.ArcGIS.DataManagementTools.Delete();
            gp.AddOutputsToMap = false;
            d.in_data = DKpath;
            gp.Execute(d, null);


            Invoke(new Action(() => { progressBar1.Value = 80; }));


            //还原地块
            ESRI.ArcGIS.DataManagementTools.Copy c = new ESRI.ArcGIS.DataManagementTools.Copy();
            gp.AddOutputsToMap = false;
            c.in_data = copy.out_data;
            c.out_data = DKpath;
            gp.Execute(c, null);


            Invoke(new Action(() => { progressBar1.Value = 90; }));
            IWorkspaceFactory w = new AccessWorkspaceFactoryClass();
            IWorkspace ws = w.OpenFromFile(workspacepathname, 0);

            IFeatureClass pFC = (ws as IFeatureWorkspace).OpenFeatureClass(DKname);
            IFeatureLayer pFLayer = new FeatureLayerClass();
            pFLayer.FeatureClass = pFC;
            pFLayer.Name = pFC.AliasName;
            ILayer pLayer = pFLayer as ILayer;
            IMap pMap = ArcMap.Document.FocusMap;
            ArcMap.Document.AddLayer(pLayer);


            //删除拓扑
            ESRI.ArcGIS.DataManagementTools.Delete dl = new ESRI.ArcGIS.DataManagementTools.Delete();
            gp.AddOutputsToMap = false;
            dl.in_data = ctp.in_dataset;
            gp.Execute(dl, null);
            ArcMap.Document.ActiveView.ContentsChanged();
            ArcMap.Document.ActiveView.Refresh();
            Invoke(new Action(() => { progressBar1.Value = 100; }));

        }
        List<ILayer> layersList;
        protected override void OnClosing(CancelEventArgs e)
        {
            if (isworkingflag == true)
                e.Cancel = true;
            base.OnClosing(e);
        }
        ILayer selectedLayer;

        private void button1_Click(object sender, EventArgs e)
        {
            bool b = ((selectedLayer as IDataset).Workspace as IWorkspaceEdit).IsBeingEdited();
            if (b)
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
                        isworkingflag = true;
                    }));
                    start();
                    Invoke(new Action(() =>
                    {
                        this.ControlBox = true;
                        this.button1.Enabled = true;
                        this.comboBox1.Enabled = true;
                        isworkingflag = false;
                    }));
                }));
            MessageBox.Show("完成");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedLayer = layersList[comboBox1.SelectedIndex];
            if ((selectedLayer as IFeatureLayer).FeatureClass.ShapeType == ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon )
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            layersList = new List<ILayer>();
            IEnumLayer pls = ArcMap.Document.FocusMap.Layers;
            ILayer pl;
            while ((pl = pls.Next()) != null)
            {
                layersList.Add(pl);
            }
            layersList.ForEach(new Action<ILayer>((player) => { comboBox1.Items.Add(player.Name); }));
            button1.Enabled = false;
        }
    }
}
