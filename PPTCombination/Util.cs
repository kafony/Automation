using System;

namespace PPTCombination
{
    public class Util
    {
        public static float ConvertToPixel(float cm)
        {
            // 72dpi 1厘米=28.346像素，300dpi 1厘米=118.11像素。
            float ratio = 28.346f;
            return cm * ratio;
        }

        public static bool Equals(float value1, float value2, float precision = 1f)
        {
            return Math.Abs(value1 - value2) < precision;
        }
    }
}
