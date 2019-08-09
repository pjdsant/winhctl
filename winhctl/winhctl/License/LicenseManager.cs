using System;
using System.Security.Cryptography;
using System.Text;

namespace PJSIT.WinHCtl.License
{
    public class LicenseManager
    {
        private DateTime dataDaExpiracao;
        private DateTime dataExpirou;
        private DateTime dataAtual;
        private bool licensed;

        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public bool Expiracao_Sistema()
        {
            licensed = false;
            try
            {
                Microsoft.Win32.RegistryKey EncryptedKey;
                RSACryptoServiceProvider crypto = new RSACryptoServiceProvider();
                EncryptedKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\EncryptedDate");

                object value = EncryptedKey.GetValue("EncryptedDate");

                if (value == null)
                {
                    EncryptedKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\EncryptedDate");

                    string keyvalue = DateTime.Now.ToString("dd-MMM-yyyy");
                    byte[] messageBytes = Encoding.Unicode.GetBytes(keyvalue);

                    string encryptedMessage = Convert.ToBase64String(messageBytes);

                    EncryptedKey.SetValue("EncryptedDate", encryptedMessage);
                }
                else
                {
                    byte[] encryptedMessage1 = Convert.FromBase64String(value.ToString());

                    string key = System.Text.Encoding.Unicode.GetString(encryptedMessage1);

                    DateTime date = Convert.ToDateTime(key);

                    dataDaExpiracao = date.AddDays(30);

                    dataExpirou = dataDaExpiracao;

                    dataAtual = DateTime.Now;

                    if (dataAtual >= dataDaExpiracao)
                    {
                        licensed = true;
                        logger.Info("[TRUE] License is Valid to this Period");
                    }
                    else
                    {
                        licensed = false;
                        logger.Info("[FALSE] License isn't valid to this Period");
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }

            return licensed;
        }
    }
}
