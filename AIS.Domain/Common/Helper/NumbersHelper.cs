using System;

namespace AIS.Domain.Common.Helper
{
    public static class NumbersHelper
    {
        public static double ToDouble(this decimal value)
        {
            return Convert.ToDouble(value);
        }
    }
}
