using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProviderDashboards.metrics
{
    public static class Strings
    {
        public static bool Match(String metricName, String cellValue)
        {
            string newMetric = string.Empty;
            string newCell = string.Empty;
            string[] metric = metricName.Trim(' ').ToLower().Split(' ');
            string[] cell = cellValue.Trim(' ').ToLower().Split(' ');

            newMetric = String.Join(String.Empty, metric);
            newCell = String.Join(String.Empty, cell);

            if (newCell.Contains(newMetric))
                return true;

            return false;

        }
    }
}
