using System;
using System.Security.Cryptography;
using System.Text;

namespace LikeWater.WinHCtl.License
{
    public class LicenseManager
    {
        private DateTime dataDaExpiracao;
        private DateTime dataExpirou;
        private DateTime dataAtual;
        private bool licensed;

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
                    }
                    else
                    {
                        licensed = false;
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
