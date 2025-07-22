using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using Gsync.Utilities.Extensions;


namespace Gsync.Utilities.HelperClasses
{
    public static class TraceUtility
    {
        [Conditional("TRACE")]
        public static void LogMethodCallOld(params object[] callingMethodParamValues)
        {
            var sf = new StackTrace();
            //var assembliesAndMethods = Enumerable.Range(0, sf.FrameCount).Select(i => (sf.GetFrame(i).GetMethod().DeclaringType.Assembly.GetName().Name, sf.GetFrame(i).GetMethod().Name)).ToArray();
            
            int frameLevel = 0;
            MethodBase method; 
            try
            {
                method = GetFirstMethodOfMine(sf, ref frameLevel);
                if (method is null)
                {
                    frameLevel = 1;
                    method = new StackFrame(skipFrames: frameLevel).GetMethod();
                }
            }
            catch
            {
                frameLevel = 1;
                method = null;
            }

            var methodParameters = method?.GetParameters();
            string methodName = "";
            if (method is not null) 
            {
                methodName = $"{GetClassName(method)}.{method.Name}";
            }

            MethodBase methodCalledBy;
            try
            {
                methodCalledBy = GetFirstMethodOfMine(sf, ref frameLevel);
                if (methodCalledBy is null)
                {
                    frameLevel = 2;
                    methodCalledBy = new StackFrame(skipFrames: frameLevel).GetMethod();
                }
            }
            catch (Exception)
            {
                frameLevel = 2;
                methodCalledBy = null;
            }
            
            var methodCaller = $"{GetClassName(methodCalledBy)}.{methodCalledBy.Name}()" ?? "";

            // Exclude out parameters
            var methodParamsExcludingOut = methodParameters?.Where(p => !p.IsOut).ToArray();
            if (methodParamsExcludingOut.Length == callingMethodParamValues.Length)
            {
                List<string> parameterList = new List<string>();
                foreach (var parameter in methodParamsExcludingOut)
                {
                    parameterList.Add($"{parameter.Name}={callingMethodParamValues[parameter.Position]}");
                }

                logger.Info($"TRACE\t{methodCaller} -> {methodName}({string.Join(", ", parameterList)})");
            }
            else
            {
                logger.Info($"TRACE\t{methodCaller} -> {method.Name}(/* Please update to pass in all parameters */)");
            }
        }

        [Conditional("TRACE")]
        public static void LogMethodCall(params object[] callingMethodParamValues)
        {
            var message = GetMethodCallLogString(callingMethodParamValues);
            logger.Info(message);
        }

        public static string GetMethodCallLogString(params object[] callingMethodParamValues)
        {
            var st = new StackTrace();
            int level = 0;

            var method = st.GetCallerMethod(ref level);
            var methodName = method is null ? "Error getting method name" : $"{GetClassName(method)}.{method.Name}";

            var methodCalledBy = st.GetCallerMethod(ref level);
            var methodCaller = methodCalledBy is null ? "Error getting caller name" : $"{GetClassName(methodCalledBy)}.{methodCalledBy.Name}";

            var paramString = method.GetParameterString(callingMethodParamValues);
            return $"TRACE\t{methodCaller} -> {method.Name}({paramString})";
        }

        [Conditional("TRACE")]
        public static void LogMethodTrace(params object[] callingMethodParamValues)
        {
            var message = GetMethodTraceString(callingMethodParamValues);
            logger.Info(message);
        }

        public static string GetMethodTraceString(params object[] callingMethodParamValues)
        {
            var st = new StackTrace(1);
            var methods = st.GetMyMethods();
            var lastMethod = methods.Pop();
            var paramString = lastMethod is null? "method not resolved": lastMethod.GetParameterString(callingMethodParamValues);
            var lastName = $"{GetClassName(lastMethod)}.{lastMethod.Name}({paramString})";
            var methodNames = methods.Select(m => $"{GetClassName(m)}.{m.Name}({GetParameterNames(m)})").ToList();
            methodNames.Add(lastName);
            return string.Join(" -> ", methodNames);
        }

        public static string GetMethodTraceString(this StackTrace st, params object[] callingMethodParamValues)
        {
            var methods = st.GetMyMethods();
            var lastMethod = methods.Pop();
            var paramString = lastMethod is null ? "method not resolved" : lastMethod.GetParameterString(callingMethodParamValues);
            var lastName = $"{GetClassName(lastMethod)}.{lastMethod.Name}({paramString})";
            var methodNames = methods.Select(m => $"{GetClassName(m)}.{m.Name}({GetParameterNames(m)})").ToList();
            methodNames.Add(lastName);
            return string.Join(" -> ", methodNames);
        }

        private static string GetParameterNames(this MethodBase method)
        {
            var methodParameters = method?.GetParameters();
            return string.Join(", ", methodParameters.Select(p => $"{p.Name}"));
        }

        private static T Pop<T>(this List<T> list)
        {
            if (list.IsNullOrEmpty()) { return default; }
            var result = list.Last();
            list.RemoveAt(list.Count - 1);
            return result;
        }

        private static string GetParameterString(this MethodBase method, params object[] callingMethodParamValues)
        {
            var methodParameters = method?.GetParameters();
            
            // Exclude out parameters
            var methodParamsExcludingOut = methodParameters?.Where(p => !p.IsOut).ToArray();
            if (methodParamsExcludingOut.Length == callingMethodParamValues.Length)
            {
                List<string> parameterList = new List<string>();
                foreach (var parameter in methodParamsExcludingOut)
                {
                    parameterList.Add($"{parameter.Name}={callingMethodParamValues[parameter.Position]}");
                }

                return string.Join(", ", parameterList);
            }
            else
            {
                return "/* Please update to pass in all parameters */";
            }
        }

        public static ParameterInfo[] GetCallerParameters(this StackTrace st) => st.GetCallerMethod()?.GetParameters();

        public static ParameterInfo[] GetCallerParameters(this StackTrace st, ref int frameLevel) => st.GetCallerMethod(ref frameLevel)?.GetParameters();

        public static MethodBase GetCallerMethod(this StackTrace st)
        {
            int frameLevel = 1;
            return GetCallerMethod(st, ref frameLevel);
        }

        public static MethodBase GetCallerMethod(this StackTrace st, ref int frameLevel)
        {
            var nextLevel = frameLevel + 1;
            MethodBase methodCalledBy = null;
            try
            {
                methodCalledBy = GetFirstMethodOfMine(st, ref frameLevel);
                if (methodCalledBy is null)
                {
                    frameLevel = nextLevel;
                    methodCalledBy = st.GetFrame(frameLevel).GetMethod();
                }
            }
            catch (Exception)
            {
                frameLevel = nextLevel;
                methodCalledBy = null;
            }
            return methodCalledBy;
        }
        
        public static string[] GetMyMethodNames(this StackTrace trace)
        {
            return trace.GetMyMethods().Select(m => $"{GetClassName(m)}.{m.Name}").ToArray();
        }

        public static string TryGetMyTraceString(this StackTrace trace)
        {
            return trace.TryGetMyTraceString("");
        }

        public static string TryGetMyTraceString(this StackTrace trace, string altFailure)
        {
            try
            {
                var traceString = string.Join(" -> ", trace.GetMyMethodNames());
                return traceString;
            }
            catch (Exception e)
            {
                logger.Error($"Error in TryGetMyTraceString. Returning empty string. Details: {e.Message}", e);
                return altFailure;
            }
        }


        public static string GetMyTraceString(this StackTrace trace)
        {
            return string.Join(" -> ", trace.GetMyMethodNames());
        }

        private static string GetClassName(this MethodBase m)
        {
            if (m.IsStatic) { return m.Module.Name; }
            else { return m.DeclaringType.Name; }
        }

        public static Assembly GetAssembly(this MethodBase m)
        {
            if (m.IsStatic) { return m.Module.Assembly; }
            else { return m.DeclaringType.Assembly; }
        }
        
        public static List<(StackFrame Frame, MethodBase Method)> GetMyFrames(this StackTrace trace) 
        {
            List<(StackFrame Frame, MethodBase Method)> result = [];

            for (int i = 0; i < trace.FrameCount; i++)
            {
                var frame = trace.GetFrame(i);
                var method = frame.GetMethod();
                
                if (method is not null && 
                    method.Name != "MoveNext" && 
                    method.GetAssembly().IsMine())
                {                    
                    result.Add((frame, method));
                }
            }
            return result;
        }
        
        internal static bool IsMine(this Assembly assembly)
        {
            return ProjectNames.Contains(assembly.GetName().Name);
        }

        public static List<MethodBase> GetMyMethods(this StackTrace trace)
        {
            return trace.GetMyFrames().Select(f => f.Method).ToList();
        }

        private static MethodBase GetFirstMethodOfMine(StackTrace trace, ref int i)
        {
            MethodBase methodCalledBy = null;
            bool repeat = true;
            do
            {
                try
                {
                    if(++i >= trace.FrameCount) 
                    {
                        methodCalledBy = null;
                        repeat = false;
                    }
                    else
                    {
                        var m = trace.GetFrame(i).GetMethod();
                        string assemblyName;
                        
                        if (m.IsStatic) { assemblyName = m.Module.Assembly.GetName().Name; }
                        else { assemblyName = m.DeclaringType.Assembly.GetName().Name; }
                        
                        if (ProjectNames.Contains(assemblyName))
                        {
                            methodCalledBy = trace.GetFrame(i).GetMethod();
                            if (methodCalledBy.Name == "MoveNext") { methodCalledBy = null; }
                            else { repeat = false; }
                        }
                    }
                    
                }
                catch (Exception)
                {
                    methodCalledBy = null;
                    repeat = false;
                }
            } while (repeat);
            return methodCalledBy;
        }

        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private static List<string> _projectNames;
        internal static List<string> ProjectNames
        {
            get
            {
                if (_projectNames is null)
                {
                    _projectNames = new List<string> 
                    { 
                        "Tags", "ToDoModel", "ToDoModel.Test", "TaskVisualization", 
                        "TaskMaster.Test", "UtilitiesCS", "UtilitiesCS.Test", "QuickFiler", 
                        "QuickFiler.Test", "TaskVisualization.Test", "SVGControl", 
                        "SVGControl.Test", "TaskTree", "TaskMaster", "UtilitiesSwordfish.NET.General", 
                        "UtilitiesSwordfish.NET.Test", "Tags.Test" 
                    };
                }
                return _projectNames;
            }
        }
        
        
    }
    
}
