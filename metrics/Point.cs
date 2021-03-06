﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProviderDashboards.metrics
{
    public class Point
    {
        /// <summary>
        /// <para>Use this to keep track of X and Y coordinates in a 2d array</para>
        /// 
        /// </summary>
        public int X { get; set; }
        public int Y { get; set; }

        public Point(int x, int y)
        {
            this.X = x;
            this.Y = y;
        }
    }
}
