using Newtonsoft.Json;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace PER_DComp.Robo
{
    public static class Util
    {
        public static void Log(Logger logger, string mensagem)
        {
            Util.Log(logger, mensagem, HttpStatusCode.OK);
        }

        public static void Log(Logger logger, string mensagem, HttpStatusCode statusCode)
        {
            if ((int)statusCode > 400 && (int)statusCode < 600 && (int)statusCode != 404)
                logger.Error(mensagem);
            else
                logger.Info(mensagem);
        }

        public static void Log(Logger logger, string mensagem, string tipo)
        {
            LogLevel logLevel;
            try
            {
                logLevel = LogLevel.FromString(tipo);
            }
            catch
            {
                logLevel = LogLevel.Info;
            }

            logger.Log(logLevel, mensagem);
        }


        /**
         * Serializa um objeto em formato json
         **/
        public static string ToJson(object value)
        {
            var settings = new JsonSerializerSettings
            {
                ReferenceLoopHandling = ReferenceLoopHandling.Ignore
            };

            return JsonConvert.SerializeObject(value, Newtonsoft.Json.Formatting.Indented, settings);
        }
    }
}
