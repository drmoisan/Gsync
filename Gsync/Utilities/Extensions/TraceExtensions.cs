using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using Gsync.Utilities.HelperClasses;

namespace Gsync.Utilities.Extensions
{
    public static class TraceExtensions
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static MethodBase GetCallerByName(this StackTrace sf, string methodName)
        {
            if (methodName.IsNullOrEmpty() || methodName == "MoveNext")
                return null;
            
            MethodBase caller = null;
            bool repeat = true;
            int i = 0;
            do
            {
                try
                {
                    if (++i >= sf.FrameCount)
                    {
                        caller = null;
                        repeat = false;
                    }
                    else
                    {
                        var m = sf.GetFrame(i).GetMethod();
                        if (m.Name == methodName)
                        {
                            caller = m;
                            repeat = false;
                        }
                    }

                }
                catch (System.Exception)
                {
                    caller = null;
                    repeat = false;
                }
            } while (repeat);
            return caller;
        }

        public static string TryGetParameterName(this MethodBase method, int index)
        {
            try
            {
                var parameterName = method.GetParameterName(index);
                return parameterName;
            }
            catch (ArgumentOutOfRangeException e) 
            { 
                logger.Error(e.Message, e);
                return "";
            }
            catch (Exception e) 
            {
                logger.Error(e.Message, e);
                return "";
            }
        }

        public static string GetParameterName(this MethodBase method, int index)
        {
            var parameters = method.GetParameters();
            if (parameters is null || parameters.Count() == 0)
            {
                throw new ArgumentOutOfRangeException($"Cannot call {method.Name}.{nameof(GetParameterName)}({index}) because {method.Name} does not have any parameters");
            }
            else if (index < 0)
            {
                throw new ArgumentOutOfRangeException($"Cannot call {method.Name}.{nameof(GetParameterName)}({index}) because {index} is less than 0");
            }
            else if (index >= parameters.Count())
            {
                throw new ArgumentOutOfRangeException($"Cannot call {method.Name}.{nameof(GetParameterName)}({index}) because {index} is greater than the highest index {parameters.Count()-1}");
            }
            return method.GetParameters()[index].Name;
        }

        public static string[] GetParameterNames(this MethodBase method)
        {
            return method.GetParameters().Select(x => x.Name).ToArray();
        }
    }
}
