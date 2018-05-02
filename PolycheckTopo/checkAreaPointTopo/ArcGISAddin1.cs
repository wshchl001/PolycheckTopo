using System;
using System.Collections.Generic;
using System.Text;
using System.IO;


namespace PolycheckTopo
{
    public class Button3 : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public Button3()
        {
        }

        protected override void OnClick()
        {
            PolycheckTopo.checkAreaPointTopo.Form3 f = new checkAreaPointTopo.Form3();
            f.ShowDialog();
            ArcMap.Application.CurrentTool = null;
        }

        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;

        }
    }
}
