using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mono.Reflection;
using System.Reflection.Emit;
using System.Reflection;
//using Mono.Cecil.Cil;

namespace Gsync.Utilities.HelperClasses.NewtonSoft
{
    public static class MonoExtension
    {
        public static void EmitOperand(this Instruction instruction, ILGenerator gen, MethodBuilder mb)
        {
            switch (instruction.OpCode.OperandType)
            {
                case OperandType.InlineNone:
                    gen.Emit(instruction.OpCode);
                    break;
                case OperandType.InlineSwitch:
                    throw new NotImplementedException();
                    //int length = il.ReadInt32();
                    //int base_offset = il.position + (4 * length);
                    //int[] branches = new int[length];
                    //for (int i = 0; i < length; i++)
                    //    branches[i] = il.ReadInt32() + base_offset;

                    //instruction.Operand = branches;
                    //break;
                case OperandType.ShortInlineBrTarget:
                    //throw new NotImplementedException();
                    gen.Emit(instruction.OpCode, (int)instruction.Offset);
                    //instruction.Operand = (((sbyte)il.ReadByte()) + il.position);
                    break;
                case OperandType.InlineBrTarget:
                    throw new NotImplementedException();
                    //instruction.Operand = il.ReadInt32() + il.position;
                    //break;
                case OperandType.ShortInlineI:
                    if (instruction.OpCode == OpCodes.Ldc_I4_S)
                        gen.Emit(instruction.OpCode, (sbyte)instruction.Operand);
                    //instruction.Operand = (sbyte)il.ReadByte();
                    else
                        gen.Emit(instruction.OpCode, (byte)instruction.Operand);
                    //instruction.Operand = il.ReadByte();
                    break;
                case OperandType.InlineI:
                    gen.Emit(instruction.OpCode, (int)instruction.Operand);
                    //instruction.Operand = il.ReadInt32();
                    break;
                case OperandType.ShortInlineR:
                    //instruction.Operand = il.ReadSingle();
                    gen.Emit(instruction.OpCode, (float)instruction.Operand);
                    break;
                case OperandType.InlineR:
                    //instruction.Operand = il.ReadDouble();
                    gen.Emit(instruction.OpCode, (double)instruction.Operand);
                    break;
                case OperandType.InlineI8:
                    //instruction.Operand = il.ReadInt64();
                    gen.Emit(instruction.OpCode, (long)instruction.Operand);
                    break;
                case OperandType.InlineSig:
                    //instruction.Operand = module.ResolveSignature(il.ReadInt32());
                    gen.Emit(instruction.OpCode, (SignatureHelper)instruction.Operand);
                    break;
                case OperandType.InlineString:
                    //instruction.Operand = module.ResolveString(il.ReadInt32());
                    gen.Emit(instruction.OpCode, (string)instruction.Operand);
                    break;
                case OperandType.InlineTok:
                    if (instruction.Operand is Type type)
                    {
                        gen.Emit(instruction.OpCode, type);
                    }
                    else if (instruction.Operand is MethodInfo methodInfo1)
                    {
                        gen.Emit(instruction.OpCode, methodInfo1);
                    }
                    else if (instruction.Operand is FieldInfo fieldInfo)
                    {
                        gen.Emit(instruction.OpCode, fieldInfo);
                    }
                    else
                    {
                        throw new InvalidCastException("Operand is not a valid type, method, or field.");
                    }
                    break;
                case OperandType.InlineType:
                    gen.Emit(instruction.OpCode, (Type)instruction.Operand);
                    break;
                case OperandType.InlineMethod:
                    if (instruction.Operand is MethodInfo methodInfo)
                    {
                        gen.Emit(instruction.OpCode, methodInfo);
                    }
                    else if (instruction.Operand is ConstructorInfo constructorInfo)
                    {
                        gen.Emit(instruction.OpCode, constructorInfo);
                    }
                    else
                    {
                        throw new InvalidCastException("Operand is neither MethodInfo nor ConstructorInfo.");
                    }
                    //gen.Emit(instruction.OpCode, (MethodInfo)instruction.Operand);
                    break;
                case OperandType.InlineField:
                    //instruction.Operand = module.ResolveMember(il.ReadInt32(), type_arguments, method_arguments);
                    gen.Emit(instruction.OpCode, (FieldInfo)instruction.Operand);
                    break;
                case OperandType.ShortInlineVar:
                    //instruction.Operand = GetVariable(instruction, il.ReadByte());                    
                case OperandType.InlineVar:
                    if (instruction.OpCode.Name.Contains("loc"))
                    {
                        var loc = (LocalVariableInfo)instruction.Operand;
                        var localBuilder = gen.DeclareLocal(loc.LocalType, loc.IsPinned);
                        gen.Emit(instruction.OpCode, localBuilder);
                    }
                    else
                    {
                        throw new NotImplementedException();
                        //var loc = (ParameterInfo)instruction.Operand;
                        //var paramBuilder = mb.DefineParameter(loc.Position, loc.Attributes, loc.Name);
                        //gen.Emit(instruction.OpCode, paramBuilder);
                    }
                    break;
                //instruction.Operand = GetVariable(instruction, il.ReadInt16());
                //break;
                default:
                    throw new NotSupportedException();
            }
        }
    }
}
