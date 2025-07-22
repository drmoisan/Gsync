using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;
using System.Collections;
using System.Diagnostics;


namespace Gsync.Utilities.Extensions
{
    public static class WinFormsExtensions
    {
        /// <summary>
        /// Traverses all controls on a form recursively and performs an action defined by "action" parameter
        /// </summary>
        /// <example>   
        /// this.ForAllControls(c =>
        /// {
        ///     if (c.GetType() == typeof(TextBox)) 
        ///     {
        ///         c.TextChanged += C_TextChanged;
        ///     }
        /// });
        /// </example>
        /// <param name="parent">Topmost control on a form to traverse</param>
        /// <param name="action"></param>
        public static void ForAllControls(this Control parent, Action<Control> action)
        {
            foreach (Control c in parent.Controls)
            {
                ForAllControls(c, action);
            }
            action(parent);
        }

        public static void ForAllControls(this IEnumerable<Control> controls, Action<Control> action)
        {
            foreach (Control c in controls)
            {
                ForAllControls(c, action);
            }
        }


        public static void ForAllControls(this Control parent, Action<Control> action, IList<Control> except)
        {
            if (!except.Contains(parent))
            {
                foreach (Control c in parent.Controls)
                {
                    ForAllControls(c, action, except);
                }
                action(parent);
            }
        }

        public static void ForAllControls(this IEnumerable<Control> controls, Action<Control> action, IList<Control> except)
        {           
            foreach (Control c in controls)
            {
                if (!except.Contains(c))
                    ForAllControls(c, action, except);
            }                       
        }

        public static void ForAllControls(this Control.ControlCollection controls, Action<Control> action, IList<Control> except)
        {
            foreach (Control c in controls)
            {
                if (!except.Contains(c))
                    ForAllControls(c, action, except);
            }
        }

        public static void ForAllControls<T>(this Control parent, T value, Action<Control, T> action, IList<Control> except)
        {
            if (!except.Contains(parent))
            {
                foreach (Control c in parent.Controls)
                {
                    ForAllControls(c, value, action, except);
                }
                action(parent, value);
            }
        }

        public static void ForAllControls<T>(this Control parent, T value, Func<Control, T, T> function, IList<Control> except)
        {
            if (!except.Contains(parent))
            {
                T seedValue = function(parent, value);
                foreach (Control c in parent.Controls)
                {
                    ForAllControls(c, seedValue, function, except);
                }
            }
            
        }

        public static void ForAllControls<T>(this Control parent, T value, Func<Control, T, T> function)
        {
            T seedValue = function(parent, value);
            foreach (Control c in parent.Controls)
            {
                ForAllControls(c, seedValue, function);
            }
        }

        public static IEnumerable<Control> GetAllChildren(this Control root)
        {
            var stack = new Stack<Control>();
            stack.Push(root);

            while (stack.Any())
            {
                var next = stack.Pop();
                foreach (Control child in next.Controls)
                    stack.Push(child);
                yield return next;
            }
        }

        public static IEnumerable<Control> GetAllChildren(this Control root, IList<Control> except)
        {
            var stack = new Stack<Control>();
            if (!except.Contains(root))
                stack.Push(root);

            while (stack.Any())
            {
                var next = stack.Pop();
                foreach (Control child in next.Controls)
                    if (!except.Contains(child))
                        stack.Push(child);
                yield return next;
            }
        }

        public static T GetAncestor<T>(this Control control) where T: class
        {
            var parent = control.Parent;
            if (parent is T) { return parent as T; }
            else if (parent.Parent is not null) { return parent.GetAncestor<T>(); }
            else { return null; }
        }

        public static T GetAncestor<T>(this Control control, bool strict) where T : class
        {
            var result = control.GetAncestor<T>();
            if (result is not null || !strict) { return result; }
            else
            {
                throw new ArgumentOutOfRangeException(
                    $"{nameof(GetAncestor)} could not find an ancestor of type {typeof(T).Name} " +
                    $"in the control hierarchy of {control.Name}.");
            }
        }

        internal static U ThrowIfNotResolved<T, U>(Func<T, U> resolver, T arg)
        {
            U result = resolver(arg);
            if (result is not null) { return result; }
            else
            {
                var methodBase = new StackTrace().GetFrame(1).GetMethod();
                //var callerName = methodBase.Name;
                var parameters = methodBase.GetParameters();
                var message = $"{parameters[0].Name} could not find an ItemViewer "
                                + $"in the control hierarchy of {parameters[1].Name}.  " +
                                $"This method must be assigned to a {nameof(ComboBox)} " +
                                $"that is a child of an {typeof(T).Name}";
                throw new ArgumentOutOfRangeException(message);
            }
        }

        public static bool IsRegistered(this EventHandler handler, 
                                        Delegate prospectiveHandler) => 
            handler != null && 
            handler.GetInvocationList()
                   .Any(existingHandler => existingHandler == prospectiveHandler);

        public static (EventHandlerList EventHandlerList, object Object) GetEventHandlerList(this object control, string eventName)
        {
            eventName = "Event" + eventName;
            FieldInfo f1 = typeof(Control).GetField(eventName,
                               BindingFlags.Static | BindingFlags.NonPublic);

            object obj = f1.GetValue(control);
            PropertyInfo pi = control.GetType().GetProperty("Events",
                                              BindingFlags.NonPublic | BindingFlags.Instance);

            return ((EventHandlerList)pi.GetValue(control, null), obj);
        }

        public static void RemoveEventHandlers(this Control control, string eventName)
        {
            (var list, var obj) = control.GetEventHandlerList(eventName);
            list.RemoveHandler(obj, list[obj]);
        }



        #region Clone<T>

        // 3 Public Overloads of Clone<T>

        public static T Clone<T>(this T controlToClone, string name, bool deep=false)
            where T : Control
        {
            T instance = controlToClone.Clone<T>(deep);
            instance.Name = name;
            return instance;
        }

        public static T Clone<T>(this T controlToClone, bool deep=false)
            where T : Control
        {
            PropertyInfo[] controlProperties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            T clonedInstance = GetInstance<T>(typeof(T));

            var excluded = new List<string>() { "WindowTarget", "Name", "Parent", }; //"LayoutSettings"
            var properties = controlProperties.Where(p => !excluded.Contains(p.Name)).ToArray();

            properties.CopyToDestination(ref clonedInstance, controlToClone, deep, 4);

            return clonedInstance;
        }

        public static T Clone<T>(this T sourceClass, bool deep, int remainingDepth) where T : class
            
        {
            var type = sourceClass.GetType();
            var properties = type.GetProperties();

            T clonedClass = GetInstance<T>(type);

            properties.CopyToDestination(ref clonedClass, sourceClass, deep, remainingDepth);

            return clonedClass;
        }

        public static RowStyle Clone(this RowStyle sourceStyle)
        {
            if (sourceStyle == null) { throw new ArgumentNullException(); }
            return new RowStyle(sourceStyle.SizeType, sourceStyle.Height);
        }

        public static ColumnStyle Clone(this ColumnStyle sourceStyle)
        {
            if (sourceStyle == null) { throw new ArgumentNullException(); }
            return new ColumnStyle(sourceStyle.SizeType, sourceStyle.Width);
        }

        // 1 Iterative Switch Function for Clone<T>

        private static void CopyToDestination<T>(this PropertyInfo[] properties, ref T destinationClass, T sourceClass, bool deep, int remainingDepth) where T : class
        {
            foreach (var property in properties)
            {
                if (property.PropertyType == typeof(TableLayoutSettings))
                {
                    property.CopyTableLayoutSettings(ref destinationClass, sourceClass);
                }
                else if (TypesSetByProperties.Contains(property.PropertyType))
                {
                    property.CopyBySubProperties(ref destinationClass, sourceClass, deep, remainingDepth);
                }
                else if (property.CanWrite)
                {
                    if (deep) { property.DeepCopyProperty(sourceClass, remainingDepth, ref destinationClass); }
                    else { property.ShallowCopyProperty(sourceClass, ref destinationClass); }
                }
                
            }
        }
        
        // 4 Types of Node Copying
        
        private static void ShallowCopyProperty<T>(this PropertyInfo propInfo, T classToClone, ref T clonedInstance)  
        {
            propInfo.SetValue(clonedInstance, propInfo.GetValue(classToClone, null), null);   
        }

        private static void DeepCopyProperty<T>(this PropertyInfo property, T classToClone, int remainingDepth, ref T clonedInstance) 
        {
            object value = property.GetValue(classToClone);
            if ((value != null) && (value.GetType().IsClass) && (!value.GetType().IsPrimitiveLike()) && (value.GetType().FullName.StartsWith("System.") ? (remainingDepth-- > 0) : true)) 
            { 
                property.SetValue(clonedInstance, value.Clone(true, remainingDepth));
            }
            else
            {
                property.SetValue(classToClone, value);
            }    
        }

        private static void CopyBySubProperties<T>(this PropertyInfo property, ref T destinationClass, T sourceClass, bool deep, int remainingDepth) where T : class
        {
            object sourceClass2 = property.GetValue(sourceClass);
            object destinationClass2 = property.GetValue(destinationClass);

            PropertyInfo[] subProperties = sourceClass2.GetType().GetProperties();

            subProperties.CopyToDestination(ref destinationClass2, sourceClass2, deep, remainingDepth);
        }

        private static void CopyTableLayoutSettings<T>(this PropertyInfo property, ref T destinationClass, T sourceClass) 
        {
            var source = (TableLayoutSettings)property.GetValue(sourceClass); 
            var destination = (TableLayoutSettings)property.GetValue(destinationClass);
                        
            destination.ColumnCount = source.ColumnCount;
            foreach (ColumnStyle style in source.ColumnStyles)
            {
                destination.ColumnStyles.Add(style.Clone());
            };
            destination.RowCount = source.RowCount;
            foreach (RowStyle style in source.RowStyles)
            {
                destination.RowStyles.Add(style.Clone());
            };
            destination.GrowStyle = source.GrowStyle;
        }

        // Clone<T> Helpers

        private static T GetInstance<T>(Type type) where T : class
        {
            T clonedInstance;

            try
            {
                clonedInstance = (T)Activator.CreateInstance(type);
            }
            catch (MissingMethodException)
            {
                //T clonedClass = RuntimeHelpers.GetUninitializedObject(type);
                clonedInstance = (T)FormatterServices.GetUninitializedObject(type);
            }
            return clonedInstance;
        }

        internal static List<Type> TypesSetByProperties = new List<Type>() 
        { 
             
        };

        internal static bool IsPrimitiveLike(this Type type)
        {
            return type.IsPrimitive || 
                   type.IsEnum || 
                   type == typeof(string) || 
                   type == typeof(decimal) ||
                   type == typeof(DateTime);
        }

        #endregion
    }
}
