using System;

namespace LikeWater.WinHCtl.License
{
    public class DateProtect
    {
        private static readonly DateTime expiration = new DateTime(2019, 08, 09);

        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public bool DataProtect()
        {
            bool licensed = false;
            try
            {
                DateTime currentTime = DateTime.Now;

                if (currentTime.Date <= expiration.Date)
                {
                    licensed = true;
                    logger.Info("[TRUE] License is Valid to this Period");
                }
                else
                {
                    logger.Info("[FALSE] License isn't valid to this Period");
                }
            }
            catch (Exception e)
            {
                logger.Error("[EXCEPTION] Occur an exception - " + e.Message);
                throw new Exception(e.Message);
            }


            return licensed;
        }
    }
}
