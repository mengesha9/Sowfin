using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Collections;

namespace Sowfin.API
{
    public class CircularMonitor : AbstractCalculationMonitor
    {
        public ArrayList circulars = new ArrayList();
        public ArrayList Circulars { get { return circulars; } }

        public override bool OnCircular(IEnumerator circularCellsData)
        {
            CalculationCell cc = null;
            ArrayList cur = new ArrayList();
            while (circularCellsData.MoveNext())
            {
                cc = (CalculationCell)circularCellsData.Current;
                cur.Add(cc.Worksheet.Name + "!" + CellsHelper.CellIndexToName(cc.CellRow, cc.CellColumn));
            }
            circulars.Add(cur);
            return true;
        }
    }
}
