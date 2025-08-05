using Mono.Reflection;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Gsync.Utilities.ReusableTypes;
using Gsync.Utilities.Extensions;

namespace Gsync.Utilities.HelperClasses.NewtonSoft
{
    public class WrapperScoDictionary<TDerived, TKey, TValue> where TDerived : ScoDictionaryNew<TKey, TValue>
    {
        [JsonProperty("CoDictionary")]
        public ConcurrentObservableDictionary<TKey, TValue> CoDictionary { get; set; }

        [JsonProperty("RemainingObject")]
        public object RemainingObject { get; set; }

        public WrapperScoDictionary()
        {
            CoDictionary = new ConcurrentObservableDictionary<TKey, TValue>();
        }

        public TDerived ToDerived(WrapperScoDictionary<TDerived, TKey, TValue> wrapper)
        {
            CoDictionary = wrapper.CoDictionary;
            RemainingObject = wrapper.RemainingObject;
            return ToDerived();
        }

        public TDerived ToDerived()
        {
            CoDictionary.ThrowIfNull();
            RemainingObject.ThrowIfNull();

            var derivedInstance = (TDerived)Activator.CreateInstance(typeof(TDerived), true);

            // Populate dictionary values
            foreach (var kvp in CoDictionary)
            {
                derivedInstance.TryAdd(kvp.Key, kvp.Value);
            }

            var derivedType = typeof(TDerived);
            var remainingType = RemainingObject.GetType();

            // Step 1: Transfer Config via property if available
            var configProp = derivedType.GetProperty("Config", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);

            // remainingType is a dynamically emitted type, so we cannot rely on reflection to get the 
            // property info unless we happened to emit it in the way that .NET would do it.
            var configField = remainingType.GetField("_Config", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            var configValue = configField?.GetValue(RemainingObject);
            
            configProp?.SetValue(derivedInstance, configValue);

            // STEP 2: Transfer other properties that are declared on the derived type
            var derivedProps = derivedType.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            var derivedFields = derivedType.GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            
            // Track which property names were set and their status (true = property is non-default after set)
            var propertiesSet = new Dictionary<string, bool>();
            // Track which fields were set via properties to avoid overwriting them in Step 3
            var fieldsSetByProperties = new HashSet<string>();
            var before = SnapshotFields(derivedInstance, derivedFields);

            foreach (var prop in derivedProps)
            {
                if (prop.Name == "Config" || prop.Name == "ism" || !prop.CanWrite) continue;
                if (prop.GetIndexParameters().Length > 0) continue;

                var sourceProp = remainingType.GetProperty(prop.Name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (sourceProp?.CanRead == true && sourceProp.GetIndexParameters().Length == 0)
                {
                    var value = sourceProp.GetValue(RemainingObject);
                    prop.SetValue(derivedInstance, value);

                    // Try to get the field that backs this property, if any
                    var backingField = derivedType.GetField($"_{prop.Name}", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public)
                                    ?? derivedType.GetField($"<{prop.Name}>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
                    if (backingField != null)
                        fieldsSetByProperties.Add(backingField.Name);                    
                }
            }

            var after = SnapshotFields(derivedInstance, derivedFields);
            foreach (var field in derivedFields)
            {
                var oldValue = before[field.Name];
                var newValue = after[field.Name];
                if (!object.Equals(oldValue, newValue))
                    fieldsSetByProperties.Add(field.Name);
            }

            // STEP 3: Transfer additional fields if they exist on both types            
            var sourceFields = remainingType.GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            foreach (var srcField in sourceFields)
            {
                if (srcField.Name == "ism" || srcField.Name == "_ism") continue;
                if (fieldsSetByProperties.Contains(srcField.Name))
                {
                    Console.WriteLine($"SKIP Step 3 field '{srcField.Name}' (set by property in Step 2)");
                    continue; // Don't overwrite property values set in Step 2
                }

                var destField = derivedFields.FirstOrDefault(f => f.Name == srcField.Name);
                if (destField != null)
                {                    
                    var value = srcField.GetValue(RemainingObject);
                    Console.WriteLine($"SET field '{srcField.Name}' to '{value ?? "null"}'");
                    destField.SetValue(derivedInstance, value);
                }
            }

            return derivedInstance;
        }

        Dictionary<string, object> SnapshotFields(object obj, IEnumerable<FieldInfo> fields)
        {
            var dict = new Dictionary<string, object>();
            foreach (var field in fields)
                dict[field.Name] = field.GetValue(obj);
            return dict;
        }

        public WrapperScoDictionary<TDerived, TKey, TValue> ToComposition(TDerived derivedInstance)
        {
            derivedInstance.ThrowIfNull();
            CoDictionary = new ConcurrentObservableDictionary<TKey, TValue>(derivedInstance);

            Type objectType = CompileType();
            var instance = CopyTo(derivedInstance, objectType);
            RemainingObject = instance;

            return this;
        }

        public Type CompileType()
        {
            var derivedType = typeof(TDerived);
            var baseType = typeof(ConcurrentObservableDictionary<TKey, TValue>);

            var derivedProperties = derivedType
                .GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                .Where(p => 
                    p.DeclaringType != baseType && 
                    p.Name != nameof(ScoDictionaryNew<TKey, TValue>.Config) && 
                    p.Name != "ism")
                .ToArray();

            var derivedFields = derivedType
                .GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                .Where(f => f.DeclaringType != baseType && f.Name != "ism" && f.Name != "_ism")
                .ToArray();

            var tb = GetTypeBuilder();
            tb.DefineDefaultConstructor(MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.RTSpecialName);

            CreateConfigProperty(tb);

            var capturedFields = new Dictionary<string, FieldBuilder>();

            foreach (var property in derivedProperties)
            {
                if (property.CanRead && property.CanWrite)
                {
                    ReplicateProperty(tb, property, ref capturedFields);
                }
            }

            var remainingFields = derivedFields
                .Where(f => !capturedFields.ContainsKey(f.Name))
                .ToArray();

            foreach (var field in remainingFields)
            {
                tb.DefineField(field.Name, field.FieldType, field.Attributes);
            }

            return tb.CreateType();
        }

        public object CopyTo(TDerived instance, Type objectType)
        {
            var myObject = Activator.CreateInstance(objectType);
            var derivedType = typeof(TDerived);

            // Set up the config field
            objectType.GetField("_Config", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                ?.SetValue(myObject, instance.Config);

            // Get all other fields in the derived type except for ism which was captured by the _Config field
            var derivedFields = derivedType.GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                .Where(field => field.Name != "ism" && field.Name != "_ism")
                .ToArray();

            foreach (var field in derivedFields)
            {
                var fieldValue = field.GetValue(instance);
                var fieldInfo = objectType.GetField(field.Name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (fieldInfo != null)
                {
                    // Debug: Field assignment
                    if (field.Name.Contains("AdditionalField3") || field.Name.Contains("_additionalField3"))
                        Console.WriteLine($"CopyTo: Setting field '{field.Name}' to '{fieldValue}'");
                    fieldInfo.SetValue(myObject, fieldValue);
                }
            }

            // **NEW: Copy property values to the object if not already handled by a field**
            var properties = derivedType.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            foreach (var prop in properties)
            {
                if (prop.Name == "ism") continue; // Skip ism property
                if (prop.GetIndexParameters().Length > 0)
                {
                    Console.WriteLine($"Skipping indexed property '{prop.Name}'");
                    continue; // Skip indexed properties
                }
                if (prop.CanRead && prop.CanWrite)
                {
                    var value = prop.GetValue(instance);
                    var propInfo = objectType.GetProperty(prop.Name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                    if (propInfo != null && propInfo.CanWrite)
                    {
                        Console.WriteLine($"CopyTo: Setting property '{prop.Name}' to '{value}'");
                        propInfo.SetValue(myObject, value);
                    }
                }
            }

            return myObject;
        }


        private TypeBuilder GetTypeBuilder()
        {
            var typeSignature = $"{typeof(TDerived).Name}_ExDictionary";
            var assemblyName = new AssemblyName(typeSignature);
            AssemblyBuilder assemblyBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly(
                assemblyName, AssemblyBuilderAccess.Run);

            ModuleBuilder moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule");

            TypeBuilder tb = moduleBuilder.DefineType(typeSignature,
                    TypeAttributes.Public |
                    TypeAttributes.Class |
                    TypeAttributes.AutoClass |
                    TypeAttributes.AnsiClass |
                    TypeAttributes.BeforeFieldInit |
                    TypeAttributes.AutoLayout,
                    null);
            return tb;
        }

        //public void CreateConfigProperty(TypeBuilder tb)
        //{
        //    var propertyBuilder = tb.DefineProperty("Config", PropertyAttributes.None, typeof(NewSmartSerializableConfig), null);
        //    var fieldBuilder = tb.DefineField("_Config", typeof(NewSmartSerializableConfig), FieldAttributes.Private);

        //    var getMethod = tb.DefineMethod("get_Config", MethodAttributes.Public, typeof(NewSmartSerializableConfig), Type.EmptyTypes);
        //    var getIl = getMethod.GetILGenerator();
        //    getIl.Emit(OpCodes.Ldarg_0);
        //    getIl.Emit(OpCodes.Ldfld, fieldBuilder);
        //    getIl.Emit(OpCodes.Ret);
        //    propertyBuilder.SetGetMethod(getMethod);

        //    var setMethod = tb.DefineMethod("set_Config", MethodAttributes.Public, null, new[] { typeof(NewSmartSerializableConfig) });
        //    var setIl = setMethod.GetILGenerator();
        //    Label modifyProperty = setIl.DefineLabel();
        //    Label exitSet = setIl.DefineLabel();
        //    setIl.MarkLabel(modifyProperty);
        //    setIl.Emit(OpCodes.Ldarg_0);
        //    setIl.Emit(OpCodes.Ldarg_1);
        //    setIl.Emit(OpCodes.Stfld, fieldBuilder);
        //    setIl.Emit(OpCodes.Nop);
        //    setIl.MarkLabel(exitSet);
        //    setIl.Emit(OpCodes.Ret);
        //}

        public void CreateConfigProperty(TypeBuilder tb)
        {
            Type configType = typeof(NewSmartSerializableConfig);

            if (CoDictionary is ScoDictionaryNew<TKey, TValue> dict && dict.Config is not null)
            {
                configType = dict.Config.GetType();
            }

            var propertyBuilder = tb.DefineProperty("Config", PropertyAttributes.None, configType, null);
            var fieldBuilder = tb.DefineField("_Config", configType, FieldAttributes.Private);

            var getMethod = tb.DefineMethod("get_Config", MethodAttributes.Public, configType, Type.EmptyTypes);
            var getIl = getMethod.GetILGenerator();
            getIl.Emit(OpCodes.Ldarg_0);
            getIl.Emit(OpCodes.Ldfld, fieldBuilder);
            getIl.Emit(OpCodes.Ret);
            propertyBuilder.SetGetMethod(getMethod);

            var setMethod = tb.DefineMethod("set_Config", MethodAttributes.Public, null, new[] { configType });
            var setIl = setMethod.GetILGenerator();
            setIl.Emit(OpCodes.Ldarg_0);
            setIl.Emit(OpCodes.Ldarg_1);
            setIl.Emit(OpCodes.Stfld, fieldBuilder);
            setIl.Emit(OpCodes.Ret);
            propertyBuilder.SetSetMethod(setMethod);
        }



        public void ReplicateProperty(TypeBuilder tb, PropertyInfo property, ref Dictionary<string, FieldBuilder> capturedFields)
        {
            var fieldName = $"_{property.Name}";
            var fieldBuilder = tb.DefineField(fieldName, property.PropertyType, FieldAttributes.Private);
            capturedFields[property.Name] = fieldBuilder;

            var propertyBuilder = tb.DefineProperty(property.Name, PropertyAttributes.HasDefault, property.PropertyType, null);

            // Define getter method
            var getMethod = tb.DefineMethod($"get_{property.Name}",
                MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig,
                property.PropertyType, Type.EmptyTypes);

            var getIL = getMethod.GetILGenerator();
            getIL.Emit(OpCodes.Ldarg_0);
            getIL.Emit(OpCodes.Ldfld, fieldBuilder);
            getIL.Emit(OpCodes.Ret);

            propertyBuilder.SetGetMethod(getMethod);

            // Define setter method
            var setMethod = tb.DefineMethod($"set_{property.Name}",
                MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig,
                null, new[] { property.PropertyType });

            var setIL = setMethod.GetILGenerator();
            setIL.Emit(OpCodes.Ldarg_0);
            setIL.Emit(OpCodes.Ldarg_1);
            setIL.Emit(OpCodes.Stfld, fieldBuilder);
            setIL.Emit(OpCodes.Ret);

            propertyBuilder.SetSetMethod(setMethod);
        }

        //public void ReplicateProperty(TypeBuilder tb, PropertyInfo property, ref Dictionary<string, FieldBuilder> capturedFields)
        //{
        //    PropertyBuilder propertyBuilder = tb.DefineProperty(property.Name, property.Attributes, property.PropertyType, property.DeclaringType.GetGenericArguments());
        //    var getMethod = ModifyGetMethod(tb, property, ref capturedFields);
        //    if (getMethod is not null) { propertyBuilder.SetGetMethod(getMethod); }

        //    var setMethod = ModifySetMethod(tb, property, ref capturedFields);
        //    if (setMethod is not null) { propertyBuilder.SetSetMethod(setMethod); }
        //    ;
        //}

        public void ReplicateProperty(TypeBuilder tb, PropertyInfo property, FieldInfo existingField)
        {
            //FieldBuilder fieldBuilder = tb.DefineField("_" + propertyName, propertyType, FieldAttributes.Private);
            var fieldBuilder = tb.DefineField(existingField.Name, existingField.FieldType, existingField.Attributes);
            var getAttributes = property.GetGetMethod().Attributes;
            var setAttributes = property.GetSetMethod().Attributes;

            PropertyBuilder propertyBuilder = tb.DefineProperty(property.Name, property.Attributes, property.PropertyType, null);
            MethodBuilder getPropMthdBldr = GenerateGetMethod(tb, property, fieldBuilder, getAttributes);
            MethodBuilder setPropMthdBldr = GenerateSetMethod(tb, property, fieldBuilder, setAttributes);

            propertyBuilder.SetGetMethod(getPropMthdBldr);
            propertyBuilder.SetSetMethod(setPropMthdBldr);
        }

        private static MethodBuilder GenerateSetMethod(TypeBuilder tb, PropertyInfo property, FieldBuilder fieldBuilder, MethodAttributes setAttributes)
        {
            MethodBuilder setPropMthdBldr =
                            tb.DefineMethod("set_" + property.Name,
                              setAttributes,
                              null, new[] { property.PropertyType });

            ILGenerator setIl = setPropMthdBldr.GetILGenerator();
            Label modifyProperty = setIl.DefineLabel();
            Label exitSet = setIl.DefineLabel();

            setIl.MarkLabel(modifyProperty);
            setIl.Emit(OpCodes.Ldarg_0);
            setIl.Emit(OpCodes.Ldarg_1);
            setIl.Emit(OpCodes.Stfld, fieldBuilder);

            setIl.Emit(OpCodes.Nop);
            setIl.MarkLabel(exitSet);
            setIl.Emit(OpCodes.Ret);
            return setPropMthdBldr;
        }

        //private MethodBuilder ModifySetMethod(TypeBuilder tb, PropertyInfo property, ref Dictionary<string, FieldBuilder> backingFields)
        //{
        //    //Type[] method_arguments = null;
        //    Type[] type_arguments = null;
        //    var oldSetMethod = property.GetSetMethod(true);
        //    if (oldSetMethod == null) { return null; }

        //    //if (!(oldGetMethod is ConstructorInfo))
        //    //    method_arguments = oldGetMethod.GetGenericArguments();

        //    if (oldSetMethod.DeclaringType != null)
        //        type_arguments = oldSetMethod.DeclaringType.GetGenericArguments();

        //    var oldInstructions = Disassembler.GetInstructions(oldSetMethod);
        //    //var newInstructions = new List<Instruction>();

        //    MethodBuilder setPropMthdBldr = tb.DefineMethod("set_" + property.Name, oldSetMethod.Attributes, property.PropertyType, type_arguments);
        //    ILGenerator setIl = setPropMthdBldr.GetILGenerator();

        //    foreach (var instruction in oldInstructions)
        //    {
        //        if (instruction.OpCode == OpCodes.Ldfld || instruction.OpCode == OpCodes.Stfld)
        //        {
        //            var bf = (FieldInfo)instruction.Operand;
        //            //FieldBuilder fieldBuilder;
        //            if (!backingFields.TryGetValue(bf.Name, out var fieldBuilder))
        //            {
        //                fieldBuilder = tb.DefineField(bf.Name, bf.FieldType, bf.Attributes);
        //                backingFields[bf.Name] = fieldBuilder;
        //            }

        //            setIl.Emit(instruction.OpCode, fieldBuilder);
        //            //setIl.Emit(OpCodes.Ldfld, fieldBuilder);
        //        }
        //        else if (instruction.OpCode == OpCodes.Callvirt)
        //        {
        //            var method = (MethodInfo)instruction.Operand;
        //            setIl.Emit(instruction.OpCode, method);
        //        }
        //        else if (instruction.Operand is not null)
        //        {
        //            instruction.EmitOperand(setIl, setPropMthdBldr);
        //            //getIl.Emit(instruction.OpCode, instruction.Operand);
        //        }
        //        else
        //        {
        //            setIl.Emit(instruction.OpCode);
        //        }
        //        //newInstructions.Add(instruction);
        //    }

        //    return setPropMthdBldr;

        //}

        //private MethodBuilder ModifyGetMethod(TypeBuilder tb, PropertyInfo property, ref Dictionary<string, FieldBuilder> backingFields)
        //{
        //    //Type[] method_arguments = null;
        //    Type[] type_arguments = null;
        //    var oldGetMethod = property.GetGetMethod(true);
        //    if (oldGetMethod == null) { throw new InvalidOperationException("Property does not have a getter."); }

        //    //if (!(oldGetMethod is ConstructorInfo))
        //    //    method_arguments = oldGetMethod.GetGenericArguments();

        //    if (oldGetMethod.DeclaringType != null)
        //        type_arguments = oldGetMethod.DeclaringType.GetGenericArguments();

        //    var oldInstructions = Disassembler.GetInstructions(oldGetMethod);
        //    //var newInstructions = new List<Instruction>();

        //    MethodBuilder getPropMthdBldr = tb.DefineMethod("get_" + property.Name, oldGetMethod.Attributes, property.PropertyType, type_arguments);
        //    ILGenerator getIl = getPropMthdBldr.GetILGenerator();

        //    foreach (var instruction in oldInstructions)
        //    {
        //        if (instruction.OpCode == OpCodes.Ldfld || instruction.OpCode == OpCodes.Stfld)
        //        {
        //            var bf = (FieldInfo)instruction.Operand;
        //            //FieldBuilder fieldBuilder;
        //            if (!backingFields.TryGetValue(bf.Name, out var fieldBuilder))
        //            {
        //                fieldBuilder = tb.DefineField(bf.Name, bf.FieldType, bf.Attributes);
        //                backingFields[bf.Name] = fieldBuilder;
        //            }

        //            getIl.Emit(instruction.OpCode, fieldBuilder);
        //            //getIl.Emit(OpCodes.Ldfld, fieldBuilder);
        //        }
        //        else if (instruction.OpCode == OpCodes.Callvirt)
        //        {
        //            var method = (MethodInfo)instruction.Operand;
        //            getIl.Emit(instruction.OpCode, method);
        //        }
        //        else if (instruction.Operand is not null)
        //        {
        //            instruction.EmitOperand(getIl, getPropMthdBldr);
        //            //getIl.Emit(instruction.OpCode, instruction.Operand);
        //        }
        //        else
        //        {
        //            getIl.Emit(instruction.OpCode);
        //        }
        //        //newInstructions.Add(instruction);
        //    }

        //    return getPropMthdBldr;

        //}

        private static MethodBuilder GenerateGetMethod(TypeBuilder tb, PropertyInfo property, FieldBuilder fieldBuilder, MethodAttributes getAttributes)
        {
            MethodBuilder getPropMthdBldr = tb.DefineMethod("get_" + property.Name, getAttributes, property.PropertyType, Type.EmptyTypes);
            ILGenerator getIl = getPropMthdBldr.GetILGenerator();

            getIl.Emit(OpCodes.Ldarg_0);
            getIl.Emit(OpCodes.Ldfld, fieldBuilder);
            getIl.Emit(OpCodes.Ret);
            return getPropMthdBldr;
        }

        public FieldInfo GetBackingField(PropertyInfo property)
        {
            var getMethod = property.GetGetMethod(true);
            if (getMethod == null)
            {
                throw new InvalidOperationException("Property does not have a getter.");
            }

            //// New Code
            //var instructions2 = Disassembler.GetInstructions(getMethod);
            //SDILReader.MethodBodyReader reader = new SDILReader.MethodBodyReader(getMethod);
            //string bodyText = reader.GetBodyCode();
            //// End New Code

            var instructions = getMethod.GetMethodBody().GetILAsByteArray();
            for (int i = 0; i < instructions.Length; i++)
            {
                // Look for the "ldfld" or "stfld" opcode, which is used to load or store a field
                if (instructions[i] == OpCodes.Ldfld.Value || instructions[i] == OpCodes.Stfld.Value)
                {
                    // The next bytes represent the metadata token for the field
                    int metadataToken = BitConverter.ToInt32(instructions, i + 1);
                    return getMethod.Module.ResolveField(metadataToken);
                }
            }

            throw new InvalidOperationException("Backing field not found.");
        }
    }
}
