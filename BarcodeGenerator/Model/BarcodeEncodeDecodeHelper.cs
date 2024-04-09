using System;
using System.Runtime.InteropServices;
using System.Text;
using BarcodeProcessing;

namespace BarcodeGenerator.Model
{
    public abstract class BarcodeEncodeDecodeHelper : StbsBarcodeBase
    {
        private static string CurveBarcodeProcess(bool isEncode, string code)
        {
            StringBuilder sbCode = new StringBuilder();
            int count = code.Length / 16;
            string leftCode = code.Substring(0, 2);
            if (isEncode)
            {
                sbCode.Append(leftCode);
                code = code.Substring(2).Replace('.', 'a');
                for (int i = 0; i < count; i++)
                {
                    var temp = EncodeProcess(code.Substring(i * 16, 16));
                    sbCode.Append(temp);
                }
            }
            else
            {
                sbCode.Append(leftCode);
                code = code.Substring(2);
                for (int i = 0; i < count; i++)
                {
                    sbCode.Append(DecodeProcess(code.Substring(i * 16, 16)));
                }

                sbCode = sbCode.Replace('a', '.', 2, sbCode.Length - 2);
            }
            return sbCode.ToString();
        }

        public static string EncodeIntegrationBarcode(string code)
        {
            return CurveBarcodeProcess(true, code);
        }

        public static string DecodeIntegrationBarcode(string code)
        {
            return CurveBarcodeProcess(false, code);
        }

        public static string DecodeReagentBarcode(string code)
        {
            return DecodeProcess(code);
        }

        public static string EncodeReagentBarcode(string code)
        {
            return EncodeProcess(code);
        }

        public static int SuperEfficBarcodeLength = 386;//434

        public static int SuperFlexBarcodeLength = 338;
    }
}