using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gsync.Utilities.ReusableTypes
{
    public static class ScoDictionaryStatic
    {
        public static bool IsDerivedFrom_ScoDictionaryNew(this Type type)
        {
            if (type == null)
            {
                throw new ArgumentNullException(nameof(type));
            }

            Type baseType = typeof(ScoDictionaryNew<,>);
            while (type != null && type != typeof(object))
            {
                if (type.IsGenericType && type.GetGenericTypeDefinition() == baseType)
                {
                    return true;
                }
                type = type.BaseType;
            }
            return false;
        }

        public static Type[] GetScoDictionaryNewGenerics(this Type type)
        {
            if (type == null)
            {
                throw new ArgumentNullException(nameof(type));
            }

            Type baseType = typeof(ScoDictionaryNew<,>);
            while (type != null && type != typeof(object))
            {
                if (type.IsGenericType && type.GetGenericTypeDefinition() == baseType)
                {
                    return type.GetGenericArguments(); 
                }
                type = type.BaseType;
            }
            return null;
        }
    }
}
