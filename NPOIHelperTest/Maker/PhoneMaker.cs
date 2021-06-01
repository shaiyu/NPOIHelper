using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelperTest
{

    public class PhoneMaker
    {
        // @"^1[3456789]\d{9}$"
        public static string[] NormarlPrefix = new[] {
            "13","14","15","16","17","18","19",
        };

        public static string Build(Type type = Type.Normal)
        {
            string prefix = type == Type.Normal ? BuildPrefix() : BuildDisbalePrefix();
            return prefix + BuildBody();
        }

        /// <summary>
        /// 手机前两位
        /// </summary>
        /// <returns></returns>
        private static string BuildDisbalePrefix()
        {
            var prefix = new Random().Next(10, 99).ToString();

            while (NormarlPrefix.Contains(prefix))
            {
                return BuildDisbalePrefix();
            }
            return prefix;
        }

        private static string BuildPrefix()
        {
            var i = new Random().Next(0, NormarlPrefix.Length);
            return NormarlPrefix[i];
        }


        /// <summary>
        /// 手机后九位
        /// </summary>
        /// <returns></returns>
        public static string BuildBody()
        {
            var body = string.Empty;
            var rand = new Random();
            for (int i = 0; i < 9; i++)
            {
                body += rand.Next(0, 10);
            }
            return body;
        }

        /// <summary>
        /// 手机类型 0 正常 1 不可用
        /// 0  139xxxxXXXX  189xxxxXXXX
        /// 1  119xxxxXXXX  219xxxxXXXX
        /// </summary>
        public enum Type
        {
            Normal,
            Disable
        }
    }
}
