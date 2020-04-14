using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Threading;
#if USE35
using Microsoft.Scripting.Ast;
#else
using System.Linq.Expressions;
#endif

namespace Dlrsoft.VBScript.Runtime
{
    public class HelperFunctions
    {
        public static object GetDefaultPropertyValue(object target)
        {
            if (target == null) return target;

            while (target.GetType().IsCOMObject)
            {
                target = target.GetType().InvokeMember(string.Empty,
                    BindingFlags.Default | BindingFlags.Public | BindingFlags.IgnoreCase | BindingFlags.Instance | BindingFlags.GetProperty,
                    null,
                    target,
                    null);
            }

            return target;
        }

        public static object Concatenate(object left, object right)
        {
            bool leftDBNull = false;

            if (left == null)
            {
                left = String.Empty;
            }
            else if (left is DBNull)
            {
                left = String.Empty;
                leftDBNull = true;
            }
            else 
            {
                if (left.GetType().IsCOMObject)
                {
                    left = GetDefaultPropertyValue(left);
                }

                if (!(left is string))
                {
                    left = left.ToString();
                }
            }

            if (right == null)
            {
                right = string.Empty;
            }
            else if (right is DBNull)
            {
                if (leftDBNull) return DBNull.Value;
                right = string.Empty;
            }
            else 
            {
                if (right.GetType().IsCOMObject)
                {
                    right = GetDefaultPropertyValue(right);
                }

                if (!(right is string))
                {
                    right = right.ToString();
                }
            }

            return (string)left + (string)right;
        }

       

        public static object BinaryOp(ExpressionType op, object left, object right)
        {
            if (left == null || right == null)
            {
                if (op == ExpressionType.Equal)
                {
                    if (left == null && right == null)
                    {
                        return true;
                    }
                    else if (string.Empty.Equals(left) || string.Empty.Equals(right))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else if (op == ExpressionType.NotEqual)
                {
                    if (left == null && right == null)
                    {
                        return false;
                    }
                    else if (string.Empty.Equals(left) || string.Empty.Equals(right))
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    return null;
                }
            }

            Type ltype = left.GetType();
            Type rtype = right.GetType();

            if (op == ExpressionType.Equal)
            {
                if (left == right) return true;
            }


            if (ltype.IsCOMObject)
            {
                left = GetDefaultPropertyValue(left);
                ltype = left.GetType();
            }

            if (rtype.IsCOMObject)
            {
                right = GetDefaultPropertyValue(right);
                rtype = right.GetType();
            }

            if (!ltype.IsValueType && ltype != typeof(string))
            {
                left = left.ToString();
                ltype = typeof(string);
            }

            if (!rtype.IsValueType && rtype != typeof(string))
            {
                right = right.ToString();
                rtype = typeof(string);
            }

            Type targetType;
            if (ltype == rtype || CanConvert(rtype, ltype))
            {
                targetType = ltype;
            }
            else
            {
                targetType = rtype;
            }

            switch (op)
            {
                case ExpressionType.Add:
                    if (targetType == typeof(int))
                    {
                        return Convert.ToInt32(left) + Convert.ToInt32(right);
                    } 
                    else if (targetType == typeof(double))
                    {
                        return Convert.ToDouble(left) + Convert.ToDouble(right);
                    }
                    else if (targetType == typeof(long))
                    {
                        return Convert.ToInt64(left) + Convert.ToInt64(right);
                    }
                    else if (targetType == typeof(float))
                    {
                        return Convert.ToSingle(left) + Convert.ToSingle(right);
                    }
                    else if (targetType == typeof(decimal))
                    {
                        return Convert.ToDecimal(left) + Convert.ToDecimal(right);
                    }
                    else if (targetType == typeof(string))
                    {
                        return (string)left + (string)right;
                    }
                    break;
                case ExpressionType.Subtract:
                    if (targetType == typeof(int))
                    {
                        return Convert.ToInt32(left) - Convert.ToInt32(right);
                    }
                    else if (targetType == typeof(double) || targetType == typeof(string))
                    {
                        return Convert.ToDouble(left) - Convert.ToDouble(right);
                    }
                    else if (targetType == typeof(long))
                    {
                        return Convert.ToInt64(left) - Convert.ToInt64(right);
                    }
                    else if (targetType == typeof(float))
                    {
                        return Convert.ToSingle(left) - Convert.ToSingle(right);
                    }
                    else if (targetType == typeof(decimal))
                    {
                        return Convert.ToDecimal(left) - Convert.ToDecimal(right);
                    }
                    break;
                case ExpressionType.Multiply:
                    if (targetType == typeof(int))
                    {
                        return Convert.ToInt32(left) * Convert.ToInt32(right);
                    }
                    else if (targetType == typeof(double) || targetType == typeof(string))
                    {
                        return Convert.ToDouble(left) * Convert.ToDouble(right);
                    }
                    else if (targetType == typeof(long))
                    {
                        return Convert.ToInt64(left) * Convert.ToInt64(right);
                    }
                    else if (targetType == typeof(float))
                    {
                        return Convert.ToSingle(left) * Convert.ToSingle(right);
                    }
                    else if (targetType == typeof(decimal))
                    {
                        return Convert.ToDecimal(left) * Convert.ToDecimal(right);
                    }
                    break;
                case ExpressionType.Divide:
                    if (targetType == typeof(int))
                    {
                        return Convert.ToInt32(left) / Convert.ToInt32(right);
                    }
                    else if (targetType == typeof(double) || targetType == typeof(string))
                    {
                        return Convert.ToDouble(left) / Convert.ToDouble(right);
                    }
                    else if (targetType == typeof(long))
                    {
                        return Convert.ToInt64(left) / Convert.ToInt64(right);
                    }
                    else if (targetType == typeof(float))
                    {
                        return Convert.ToSingle(left) / Convert.ToSingle(right);
                    }
                    else if (targetType == typeof(decimal))
                    {
                        return Convert.ToDecimal(left) / Convert.ToDecimal(right);
                    }
                    break;
                case ExpressionType.Equal:
                    try
                    {
                        if (targetType == typeof(int))
                        {
                            return Convert.ToInt32(left) == Convert.ToInt32(right);
                        }
                        else if (targetType == typeof(double))
                        {
                            return Convert.ToDouble(left) == Convert.ToDouble(right);
                        }
                        else if (targetType == typeof(long))
                        {
                            return Convert.ToInt64(left) == Convert.ToInt64(right);
                        }
                        else if (targetType == typeof(float))
                        {
                            return Convert.ToSingle(left) == Convert.ToSingle(right);
                        }
                        else if (targetType == typeof(decimal))
                        {
                            return Convert.ToDecimal(left) == Convert.ToDecimal(right);
                        }
                        else if (targetType == typeof(string))
                        {
                            return ((string)left).Equals(right);
                        }
                        else
                        {
                            return left == right;
                        }
                    }
                    catch (Exception ex)
                    {
                        return false;
                    }
                case ExpressionType.NotEqual:
                    try
                    {
                        if (targetType == typeof(int))
                        {
                            return Convert.ToInt32(left) != Convert.ToInt32(right);
                        }
                        else if (targetType == typeof(double))
                        {
                            return Convert.ToDouble(left) != Convert.ToDouble(right);
                        }
                        else if (targetType == typeof(long))
                        {
                            return Convert.ToInt64(left) != Convert.ToInt64(right);
                        }
                        else if (targetType == typeof(float))
                        {
                            return Convert.ToSingle(left) != Convert.ToSingle(right);
                        }
                        else if (targetType == typeof(decimal))
                        {
                            return Convert.ToDecimal(left) != Convert.ToDecimal(right);
                        }
                        else if (targetType == typeof(string))
                        {
                            return !((string)left).Equals(right);
                        }
                        else
                        {
                            return left != right;
                        }
                    }
                    catch (Exception ex)
                    {
                        return true;
                    }

                case ExpressionType.GreaterThan:
                    if (targetType == typeof(int))
                    {
                        return Convert.ToInt32(left) > Convert.ToInt32(right);
                    }
                    else if (targetType == typeof(double))
                    {
                        return Convert.ToDouble(left) > Convert.ToDouble(right);
                    }
                    else if (targetType == typeof(long))
                    {
                        return Convert.ToInt64(left) > Convert.ToInt64(right);
                    }
                    else if (targetType == typeof(float))
                    {
                        return Convert.ToSingle(left) > Convert.ToSingle(right);
                    }
                    else if (targetType == typeof(decimal))
                    {
                        return Convert.ToDecimal(left) > Convert.ToDecimal(right);
                    }
                    else if (targetType == typeof(string))
                    {
                        int result = ((string)left).CompareTo(right);
                        return (result > 0);
                    }
                    break;
                case ExpressionType.GreaterThanOrEqual:
                    if (targetType == typeof(int))
                    {
                        return Convert.ToInt32(left) >= Convert.ToInt32(right);
                    }
                    else if (targetType == typeof(double))
                    {
                        return Convert.ToDouble(left) >= Convert.ToDouble(right);
                    }
                    else if (targetType == typeof(long))
                    {
                        return Convert.ToInt64(left) >= Convert.ToInt64(right);
                    }
                    else if (targetType == typeof(float))
                    {
                        return Convert.ToSingle(left) >= Convert.ToSingle(right);
                    }
                    else if (targetType == typeof(decimal))
                    {
                        return Convert.ToDecimal(left) >= Convert.ToDecimal(right);
                    }
                    else if (targetType == typeof(string))
                    {
                        int result = ((string)left).CompareTo(right);
                        return (result >= 0);
                    }
                    break;
                case ExpressionType.LessThan:
                    if (targetType == typeof(int))
                    {
                        return Convert.ToInt32(left) < Convert.ToInt32(right);
                    }
                    else if (targetType == typeof(double))
                    {
                        return Convert.ToDouble(left) < Convert.ToDouble(right);
                    }
                    else if (targetType == typeof(long))
                    {
                        return Convert.ToInt64(left) < Convert.ToInt64(right);
                    }
                    else if (targetType == typeof(float))
                    {
                        return Convert.ToSingle(left) < Convert.ToSingle(right);
                    }
                    else if (targetType == typeof(decimal))
                    {
                        return Convert.ToDecimal(left) < Convert.ToDecimal(right);
                    }
                    else if (targetType == typeof(string))
                    {
                        int result = ((string)left).CompareTo(right);
                        return (result < 0);
                    }
                    break;
                case ExpressionType.LessThanOrEqual:
                    if (targetType == typeof(int))
                    {
                        return Convert.ToInt32(left) <= Convert.ToInt32(right);
                    }
                    else if (targetType == typeof(double))
                    {
                        return Convert.ToDouble(left) <= Convert.ToDouble(right);
                    }
                    else if (targetType == typeof(long))
                    {
                        return Convert.ToInt64(left) <= Convert.ToInt64(right);
                    }
                    else if (targetType == typeof(float))
                    {
                        return Convert.ToSingle(left) <= Convert.ToSingle(right);
                    }
                    else if (targetType == typeof(decimal))
                    {
                        return Convert.ToDecimal(left) <= Convert.ToDecimal(right);
                    }
                    else if (targetType == typeof(string))
                    {
                        int result = ((string)left).CompareTo(right);
                        return (result <= 0);
                    }
                    break;
                case ExpressionType.And:
                    if (targetType == typeof(bool))
                    {
                        return (bool)left && (bool)right;
                    }
                    else if (targetType == typeof(int))
                    {
                        return Convert.ToInt32(left) & Convert.ToInt32(right);
                    }
                    else if (targetType == typeof(long))
                    {
                        return Convert.ToInt64(left) & Convert.ToInt64(right);
                    }
                    else
                    {
                        return (bool)BuiltInFunctions.CBool(left) && (bool)BuiltInFunctions.CBool(right);
                    }
                    break;
                case ExpressionType.Or:
                    if (targetType == typeof(bool))
                    {
                        return (bool)left || (bool)right;
                    }
                    else if (targetType == typeof(int))
                    {
                        return Convert.ToInt32(left) | Convert.ToInt32(right);
                    }
                    else if (targetType == typeof(long))
                    {
                        return Convert.ToInt64(left) | Convert.ToInt64(right);
                    }
                    else
                    {
                        return (bool)BuiltInFunctions.CBool(left) || (bool)BuiltInFunctions.CBool(right);
                    }
                    break;
            }
            throw new ArgumentException(string.Format("Operation {0} between {1} and {2} is not implemeted.", op, ltype.Name, rtype.Name));
        }

        private static bool CanConvert(Type fromType, Type toType)
        {
            if (toType.IsAssignableFrom(fromType)) return true;

            if (fromType == typeof(string) && (toType.IsPrimitive))
            {
                return true;
            }

            return false;
        }

        public static object Redim(Type t, params int[] ubounds)
        {
            return redimInternal(t, ubounds);
        }

        public static Array redimInternal(Type t, params int[] ubounds)
        {
            if (ubounds == null) throw new ArgumentException("Must supply bound(s) when redim an array.");
            for (int i = 0; i < ubounds.Length; i++)
            {
                //Convert from ubound to length
                ubounds[i] += 1;
            }


            return Array.CreateInstance(t, ubounds);
        }

        //public static object RedimPreserve(object array, int length1)
        //{
        //    return redimPreserveInternal(array, length1);
        //}

        //public static object RedimPreserve(object array, int length1, int length2)
        //{
        //    return redimPreserveInternal(array, length1, length2);
        //}

        //public static object RedimPreserve(object array, int length1, int length2, int length3)
        //{
        //    return redimPreserveInternal(array, length1, length2, length3);
        //}

        //public static object RedimPreserve(object array, int length1, int length2, int length3, int length4)
        //{
        //    return redimPreserveInternal(array, length1, length2, length3, length4);
        //}

        //public static object RedimPreserve(object array, int length1, int length2, int length3, int length4, int length5)
        //{
        //    return redimPreserveInternal(array, length1, length2, length3, length4, length5);
        //}

        //public static object RedimPreserve(object array, int length1, int length2, int length3, int length4, int length5, int length6)
        //{
        //    return redimPreserveInternal(array, length1, length2, length3, length4, length5, length6);
        //}

        public static object RedimPreserve(object array, params int[] lengths)
        {
            return redimPreserveInternal(array, lengths);
        }

        private static Array redimPreserveInternal(object array, params int[] lengths)
        {
            if (array == null) throw new ArgumentException("To redim an array, the array must be a previously declared array.");

            Type t = array.GetType();

            if (!t.IsArray) throw new ArgumentException("Redim variable is not an array.");

            int oldDimension = ((Array)array).Rank;
            int newDimension = lengths.Length;

            if (oldDimension != newDimension) throw new ArgumentException(string.Format("Cannot change dimension of the array from {0} to {1}", oldDimension, newDimension));

            int elementsToCopy = 1;
            for (int i = 0; i < oldDimension - 1; i++)
            {
                int oldBound = ((Array)array).GetUpperBound(i);
                if (oldBound != lengths[i])
                    throw new ArgumentException(string.Format("Can only resize the last dimension. You are trying to resize dimension {0} from {1} to {2}",
                        i + 1, oldBound, lengths[i]));
            
                elementsToCopy *= ((Array)array).GetLength(i);
            }

            elementsToCopy *= (((Array)array).GetLength(oldDimension - 1) < lengths[oldDimension - 1]) ? ((Array)array).GetLength(oldDimension - 1) : lengths[oldDimension - 1];

            Array newArray = redimInternal(t.GetElementType(), lengths);

            Array.Copy((Array)array, newArray, elementsToCopy);

            return newArray;
        }

        public static void SetError(ErrObject err, Exception ex)
        {
            //Don't catch ThreadAbortException since it is captured by Response.Redirect
            if (ex is ThreadAbortException)
                throw ex;

            err.internalRaise(ex);
        }

        public static object Eqv(object left, object right)
        {
            if (left is DBNull || right is DBNull) return DBNull.Value;

            CheckLogicOperand(ref left);
            CheckLogicOperand(ref right);
            if (left is int) BoolToInt(ref right);
            if (right is int) BoolToInt(ref left);

            if (left is bool)
                return (bool)left == (bool)right;

            return ~((int)left ^ (int)right);
        }

        public static object Imp(object left, object right)
        {
            throw new NotImplementedException("Imp method is not implemented.");
        }

        public static object Not(object target)
        {
            if (target == null) return null;

            if (target.GetType().IsCOMObject)
            {
                target = GetDefaultPropertyValue(target);
            }

            return !(bool)BuiltInFunctions.CBool(target);
        }

        public static object Negate(object target)
        {
            if (target == null) return null;

            if (target.GetType().IsCOMObject)
            {
                target = GetDefaultPropertyValue(target);
            }

            return -(double)BuiltInFunctions.CDbl(target);
        }

        private static void CheckLogicOperand(ref object value)
        {
            if (value == null) value = 0;

            if (value is float || value is double)
            {
                value = Convert.ToInt32(value);
            }

            if (value is string)
            {
                bool boolResult;
                int intResult;
                if (bool.TryParse((string)value, out boolResult))
                {
                    value = boolResult;
                }
                else if (int.TryParse((string)value, out intResult))
                {
                    value = intResult;
                }
            }

            if (value is bool || value is int) return;

            throw new ArgumentException("Logic functions or operators requires arguments that can be converted to bool or int.");
        }

        private static void BoolToInt(ref object value)
        {
            if (value is bool) value = (bool)value ? -1 : 0;
        }
    }
}
