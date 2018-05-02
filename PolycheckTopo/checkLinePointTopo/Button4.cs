using System;
using System.Collections.Generic;
using System.Text;
using System.IO;


namespace PolycheckTopo
{
    public class Button4 : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public Button4()
        {
        }

        protected override void OnClick()
        {
            PolycheckTopo.checkLinePointTopo.Form4 f = new checkLinePointTopo.Form4();
            f.ShowDialog();
            ArcMap.Application.CurrentTool = null;
        }

        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }
    }
}
