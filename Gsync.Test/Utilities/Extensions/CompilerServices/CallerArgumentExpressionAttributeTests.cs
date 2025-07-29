using System;
using Gsync.Utilities.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Runtime.CompilerServices;

namespace Gsync.Test.Utilities.Extensions.CompilerServices
{
    [TestClass]
    public class CallerArgumentExpressionAttributeTests
    {
        [TestMethod]
        public void Constructor_SetsParameterName()
        {
            // Arrange
            var expected = "paramName";

            // Act
            var attr = new CallerArgumentExpressionAttribute(expected);

            // Assert
            Assert.AreEqual(expected, attr.ParameterName);
        }

        [TestMethod]
        public void AttributeUsage_IsParameterOnly_AndNotInherited_AndNotAllowMultiple()
        {
            // Arrange
            var attrType = typeof(CallerArgumentExpressionAttribute);
            var usage = (AttributeUsageAttribute)Attribute.GetCustomAttribute(attrType, typeof(AttributeUsageAttribute));

            // Assert
            Assert.IsNotNull(usage);
            Assert.AreEqual(AttributeTargets.Parameter, usage.ValidOn);
            Assert.IsFalse(usage.Inherited);
            Assert.IsFalse(usage.AllowMultiple);
        }
    }
}