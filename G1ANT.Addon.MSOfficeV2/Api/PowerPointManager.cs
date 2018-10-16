/**
* Copyright(C) G1ANT Ltd, All rights reserved
* Solution G1ANT.Addon.MSOfficeV2, Project G1ANT.Addon.MSOfficeV2
* www.g1ant.com
*
* Licensed under the G1ANT license.
* See License.txt file in the project root for full license information.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;

namespace G1ANT.Addon.MSOfficeV2
{
    public static class PowerPointManager
	{
		private static List<PowerPointWrapper> launchedPowerPoints = new List<PowerPointWrapper>();

		public static PowerPointWrapper CurrentPowerPoint { get; private set; }

		public static PowerPointWrapper AddPowerPoint()
		{
			PowerPointWrapper wrapper = new PowerPointWrapper();
			launchedPowerPoints.Add(wrapper);
			CurrentPowerPoint = wrapper;
            return wrapper;
        }

		internal static int GetNextId()
		{
			return launchedPowerPoints.Count() > 0 ? launchedPowerPoints.Max(x => x.Id) + 1 : 0;
		}

		internal static bool Switch(int id)
		{
			PowerPointWrapper ww = launchedPowerPoints.Where(x => x.Id == id).FirstOrDefault();
			CurrentPowerPoint = ww ?? CurrentPowerPoint;
			CurrentPowerPoint.Show();
            return ww != null;
        }

		public static void Remove(PowerPointWrapper powerPointWrapper)
		{
            launchedPowerPoints.Remove(powerPointWrapper);
		}
	}
}
