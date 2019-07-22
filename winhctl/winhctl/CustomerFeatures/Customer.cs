using LikeWater.WinHCtl.WinApi;
using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace LikeWater.WinHCtl.CustomerFeatures
{

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("LikeWater.WinHandlerControlEx")]
    [ComVisible(true)]

    public class Customer
    {
        //Pegar os telefones em duas janelas
        //Formatar os campos 
        //Devolver no formato 1111111111;111111111;1111111111;111111111

        public string GetPhones(string winTitle, int idxPhoneOne, int idxPhoneTwo, string chwinTitle, int idxPhoneThree)
        {
            var retPhones = "";
            try
            {

                WinApiX winApi = new WinApiX();

                var firstPhone = winApi.GetText(winTitle, idxPhoneOne);
                var secondPhone = winApi.GetText(winTitle, idxPhoneTwo);
                var thirdPhone = winApi.GetText(chwinTitle, idxPhoneThree);

                Regex.Replace(firstPhone, "[^0-9]", "");
                Regex.Replace(secondPhone, "[^0-9]", "");
                Regex.Replace(thirdPhone, "[^0-9]", "");

                retPhones = Regex.Replace(firstPhone, "[^0-9]", "") +
                    "," + Regex.Replace(secondPhone, "[^0-9]", "") +
                    "," + Regex.Replace(thirdPhone, "[^0-9]", "");
            }
            catch (System.Exception)
            {
                //Todo Gerar Log
                throw;
            }
            
            return retPhones;

        }

    }
}
