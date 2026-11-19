using System;

namespace MBRAuthoring.Utilities
{
    class GenericUtility
    {
        public static string GetCurrentUserNTID()
        {
            string currentUser = "Anonymous";
            try
            {
                currentUser = Environment.UserName; // Case-sensitive
                //currentUser = "BASKAG03";                
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return currentUser;
        }

        public static string IsNullAs(string string1, string string2)
        {
            return $"ISNULL({ string1},{ string2}) AS { string1}";
        }
        public static string Cma(string strIn)
        {
            return $", { strIn}";
        }
        public static string DecryptText(string encodedText)
        {
            try
            {
                Decrypt.EncDec Obj = new Decrypt.EncDec();
                string decodedText = Obj.Decrypt(encodedText);

                //use Obj.Encrypt(decodedText) - To encrypt

                return decodedText;
            }
            catch (Exception ex)
            {
                throw ex; // need changes for logging the issue in DB
            }

        }

    }
}
