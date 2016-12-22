using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;

using Gem = Autodesk.AutoCAD.Geometry;

namespace Opening_testLevel
{
    class Opening
    {
        public string otm_n = "";
        public Double shirina = 0;
        public string visota = "";
        public Double glubina = 0;

        public double Rotation = 0;
        public Gem.Scale3d ScaleFactor = new Gem.Scale3d();

        public Gem.Point3d insertPoint = new Gem.Point3d();

        public Gem.Point3d pCenter = new Gem.Point3d();
        public double MarkerRadius = 0;

        //private Double otm_v = 0;


        public double Otm_v
        {
            get
            {
                Double otm = 0;
                Double.TryParse(otm_n.Replace(',', '.'), out otm);

                Double h = 0;
                Double.TryParse(visota.Replace(',', '.'), out h);

                //Math.Round(otm + h / 1000, 3);
                return Math.Round(otm + h / 1000, 3);
            }
        }


        public double Otm_Niza
        {
            get
            {
                Double otm = 0;
                Double.TryParse(otm_n.Replace(',', '.'), out otm);
                return otm;
            }
        }

        public double OpeningHight
        {
            get
            {
                Double h = 0;
                Double.TryParse(visota.Replace(',', '.'), out h);
                return h;
            }
        }



    }

}
