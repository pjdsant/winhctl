using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LikeWater.WinHCtl.Resource
{
    public class RegistryManager
    {
        public string ReadRegistry(string strKey)
        {
            var result = "";

            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\PJSIT");

            if (key != null)
            {
                result = key.GetValue("Clicks").ToString();
            }

            return result;
        }
    }
}
