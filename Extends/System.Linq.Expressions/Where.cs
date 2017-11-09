using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using System.Reflection;

namespace System.Linq.Expressions
{
    /// <summary>
    /// Where条件扩展
    /// 用于分步实现不确定条件个数的逻辑运算组织   
    /// </summary>
    public static class Where
    {
        /// <summary>
        /// 参数替换对象
        /// </summary>
        private class ParameterReplacer : ExpressionVisitor
        {
            /// <summary>
            /// 表达式的参数
            /// </summary>
            public ParameterExpression ParameterExpression { get; private set; }

            /// <summary>
            /// 参数替换对象
            /// </summary>
            /// <param name="exp">表达式的参数</param>
            public ParameterReplacer(ParameterExpression exp)
            {
                this.ParameterExpression = exp;
            }

            /// <summary>
            /// 将表达式调度到此类中更专用的访问方法之一
            /// </summary>
            /// <param name="exp">表达式</param>
            /// <returns></returns>
            public Expression Replace(Expression exp)
            {
                return this.Visit(exp);
            }

            /// <summary>
            /// 获取表达式的参数
            /// </summary>
            /// <param name="p"></param>
            /// <returns></returns>
            protected override Expression VisitParameter(ParameterExpression p)
            {
                return this.ParameterExpression;
            }
        }

        /// <summary>
        /// 返回默认为True的条件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static Expression<Func<T, bool>> True<T>()
        {
            return item => true;
        }

        /// <summary>
        /// 返回默认为False的条件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static Expression<Func<T, bool>> False<T>()
        {
            return item => false;
        }


        /// <summary>
        /// 与逻辑运算
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="expLeft">表达式1</param>
        /// <param name="expRight">表达式2</param>
        /// <returns></returns>
        public static Expression<Func<T, bool>> And<T>(this Expression<Func<T, bool>> expLeft, Expression<Func<T, bool>> expRight)
        {
            var candidateExpr = Expression.Parameter(typeof(T), "item");
            var parameterReplacer = new ParameterReplacer(candidateExpr);

            var left = parameterReplacer.Replace(expLeft.Body);
            var right = parameterReplacer.Replace(expRight.Body);
            var body = Expression.AndAlso(left, right);

            return Expression.Lambda<Func<T, bool>>(body, candidateExpr);
        }

        /// <summary>
        /// 将数组转换为Or的表示式合集
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <typeparam name="TKey">键类型</typeparam>
        /// <param name="keySelector">键选择</param>
        /// <param name="values">包含的值</param>
        /// <returns></returns>
        private static Expression<Func<T, bool>> ConverToOrExpressions<T, TKey>(Expression<Func<T, TKey>> keySelector, IEnumerable<TKey> values)
        {
            var p = keySelector.Parameters.FirstOrDefault();
            var equals = values.Select(value => (Expression)Expression.Equal(keySelector.Body, Expression.Constant(value, typeof(TKey))));
            var body = equals.Aggregate<Expression>((accumulate, equal) => Expression.Or(accumulate, equal));
            return Expression.Lambda<Func<T, bool>>(body, p);
        }

        /// <summary>
        /// in条件
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <typeparam name="TKey">键类型</typeparam>
        /// <param name="source">源</param>
        /// <param name="keySelector">键选择</param>
        /// <param name="values">包含的值</param>
        /// <returns></returns>
        public static Expression<Func<T, bool>> AndContains<T, TKey>(this Expression<Func<T, bool>> expLeft, Expression<Func<T, TKey>> keySelector, IEnumerable<TKey> values)
        {
            if (values == null || values.Any() == false)
            {
                return expLeft;
            }
            var right = ConverToOrExpressions(keySelector, values);
            return expLeft.And(right);
        }


        /// <summary>
        /// 或逻辑运算
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="expLeft">表达式1</param>
        /// <param name="expRight">表达式2</param>
        /// <returns></returns>
        public static Expression<Func<T, bool>> Or<T>(this Expression<Func<T, bool>> expLeft, Expression<Func<T, bool>> expRight)
        {
            var candidateExpr = Expression.Parameter(typeof(T), "item");
            var parameterReplacer = new ParameterReplacer(candidateExpr);

            var left = parameterReplacer.Replace(expLeft.Body);
            var right = parameterReplacer.Replace(expRight.Body);
            var body = Expression.OrElse(left, right);

            return Expression.Lambda<Func<T, bool>>(body, candidateExpr);
        }


        /// <summary>
        /// 表达式参数类型转换
        /// </summary>
        /// <typeparam name="TNew">新类型</typeparam>
        /// <param name="expression">表达式</param>
        /// <returns></returns>
        public static Expression<Func<TNew, bool>> Cast<TNew>(this LambdaExpression expression)
        {
            var candidateExpr = Expression.Parameter(typeof(TNew), "item");
            var parameterReplacer = new ParameterReplacer(candidateExpr);

            var body = parameterReplacer.Replace(expression.Body);
            return Expression.Lambda<Func<TNew, bool>>(body, candidateExpr);
        }


        /// <summary>
        /// 与逻辑运算
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="expLeft">表达式1</param>
        /// <param name="fieldName">键2</param>
        /// <param name="value">值</param>
        /// <param name="op">操作符</param>
        /// <returns></returns>
        public static Expression<Func<T, bool>> And<T>(this Expression<Func<T, bool>> expLeft, string fieldName, object value, Operator op)
        {
            var expRight = Where.GetPredicate<T>(fieldName, value, op);
            return expLeft.And(expRight);
        }

        /// <summary>
        /// 获取表达式
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fieldName">字段名</param>
        /// <param name="value">值</param>
        /// <param name="op">操作符</param>
        /// <returns></returns>
        public static Expression<Func<T, bool>> GetPredicate<T>(string fieldName, object value, Operator op)
        {
            var field = typeof(T).GetProperty(fieldName, BindingFlags.IgnoreCase | BindingFlags.Instance | BindingFlags.Public);
            if (field == null)
            {
                throw new ArgumentException(string.Format("模型不存在字段{0}", fieldName));
            }

            var paramExp = Expression.Parameter(typeof(T), "item");
            var memberExp = Expression.MakeMemberAccess(paramExp, field);

            switch (op)
            {
                case Operator.Contains:
                case Operator.EndWith:
                case Operator.StartsWith:
                    var method = typeof(string).GetMethod(op.ToString(), new Type[] { typeof(string) });
                    var callBody = Expression.Call(memberExp, method, Expression.Constant(value, typeof(string)));
                    return Expression.Lambda(callBody, paramExp) as Expression<Func<T, bool>>;

                default:
                    var valueType = field.PropertyType;
                    var valueExp = Expression.Constant(value, valueType);
                    var expMethod = typeof(Expression).GetMethod(op.ToString(), new Type[] { typeof(Expression), typeof(Expression) });

                    var symbolBody = expMethod.Invoke(null, new object[] { memberExp, valueExp }) as Expression;
                    return Expression.Lambda(symbolBody, paramExp) as Expression<Func<T, bool>>;
            }
        }
    }
}
