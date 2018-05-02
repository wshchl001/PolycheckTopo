using System;
using System.Collections.Generic;
using System.Text;
using System.IO;


namespace PolycheckTopo
{
    public class Button2 : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public Button2()
        {
        }

        protected override void OnClick()
        {
            PolycheckTopo.checkAreaLineTopo.Form2 f = new checkAreaLineTopo.Form2();
            f.ShowDialog();
            ArcMap.Application.CurrentTool = null;
        }

        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }
    }
}
