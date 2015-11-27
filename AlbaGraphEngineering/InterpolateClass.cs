using System;
using System.Collections.Generic;
using System.Linq;
using AlbaLibrary.Interpolation;
using System.Windows.Forms;

namespace AlbaGraphEngineering
{
    public static class InterpolateClass
    {
        public static AlbaLibrary.Data.Impedance InterpolateData(this AlbaLibrary.Data.Impedance impeData, IEnumerable<double> frequency, InterpolationAlgorithms.InterpolationType InterpolateType)
        {
            try
            {
                var interpx = frequency;
                var G = InterpolationAlgorithms.Interpolate(impeData.Frequency.Data, impeData.Conductance.Data, interpx, InterpolateType).ToArray();
                var B = InterpolationAlgorithms.Interpolate(impeData.Frequency.Data, impeData.Susceptance.Data, interpx, InterpolateType).ToArray();
                return new AlbaLibrary.Data.Impedance(new double[3][] 
                {
                    interpx.ToArray(),
                    G,
                    B
                });
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Interpolate Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public static IEnumerable<IEnumerable<double>> InterpolateYData(IEnumerable<double> frequency, IEnumerable<double> outputData, IEnumerable<double> interpx, InterpolationAlgorithms.InterpolationType InterpolationType)
        {
            try
            {
                return new double[2][] 
                {
                    interpx.ToArray(),
                    InterpolationAlgorithms.Interpolate(frequency, outputData, interpx, InterpolationType).ToArray()
                };
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Interpolate Save Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// Return the interpolation x-axis
        /// </summary>
        /// <param name="frequency"></param>
        /// <returns></returns>
        public static IEnumerable<double> ReturnInterpolationXPnts(this IEnumerable<double> xaxis, int myPnts)
        {
            try
            {
                var no_steps = myPnts - 1;
                var step = (xaxis.Max() - xaxis.Min()) / no_steps;

                return Enumerable.Range(0, no_steps + 1).Select(i => xaxis.Min() + (i * step)).ToArray();
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Return Interpolation Of X-Axis", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public static IEnumerable<double> ReturnInterpolationXStep(this IEnumerable<double> xaxis, double myStep)
        {
            try
            {
                var step = myStep * 1000d;
                int no_steps = (int)((xaxis.Max() - xaxis.Min()) / step);

                return Enumerable.Range(0, no_steps + 1).Select(i => xaxis.Min() + (i * step)).ToArray();
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Return Interpolation Of X-Axis", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// Returns the interpolation type
        /// </summary>
        /// <returns></returns>
        public static InterpolationAlgorithms.InterpolationType ReturnInterpolationType(int index)
        {
            switch (index)
            {
                default:
                    return InterpolationAlgorithms.InterpolationType.Linear;
                case 1:
                    return InterpolationAlgorithms.InterpolationType.Spline;
            }
        }
    }
}
