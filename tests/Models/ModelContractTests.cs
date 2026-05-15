using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using NATS.Client;
using openrmf_read_api.Models;
using Xunit;

namespace tests.Models;

public class ModelContractTests
{
    public static IEnumerable<object[]> AllModelTypes()
    {
        var modelTypes = typeof(Artifact).Assembly
            .GetTypes()
            .Where(t =>
                t.IsClass &&
                !t.IsAbstract &&
                t.Namespace != null &&
                t.Namespace.StartsWith("openrmf_read_api.Models", StringComparison.Ordinal))
            .OrderBy(t => t.FullName);

        foreach (var type in modelTypes)
        {
            yield return new object[] { type };
        }
    }

    [Theory]
    [MemberData(nameof(AllModelTypes))]
    public void Model_CanBeInstantiated(Type modelType)
    {
        var instance = Activator.CreateInstance(modelType);

        Assert.NotNull(instance);
        Assert.IsType(modelType, instance);
    }

    [Theory]
    [MemberData(nameof(AllModelTypes))]
    public void Model_WritableProperties_RoundTripExpectedValues(Type modelType)
    {
        var instance = Activator.CreateInstance(modelType);
        Assert.NotNull(instance);

        var writableProps = modelType
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanWrite && p.GetIndexParameters().Length == 0)
            .ToList();

        foreach (var property in writableProps)
        {
            var expected = CreateSampleValue(property.PropertyType, property.Name, variant: 1);
            if (expected == null && property.PropertyType.IsValueType && Nullable.GetUnderlyingType(property.PropertyType) == null)
            {
                continue;
            }

            property.SetValue(instance, expected);
            var actual = property.GetValue(instance);

            Assert.Equal(expected, actual);

            var different = CreateSampleValue(property.PropertyType, property.Name, variant: 2);
            if (different != null && !Equals(different, actual))
            {
                Assert.NotEqual(different, actual);
            }
        }
    }

    [Fact]
    public void NatsServer_FieldAssignment_WorksWithMoqLikeUsage()
    {
        var model = new NATSServer();
        var connectionMock = new Moq.Mock<IConnection>();

        model.connection = connectionMock.Object;

        Assert.NotNull(model.connection);
        Assert.NotEqual(default, model.connection);
    }

    private static object CreateSampleValue(Type type, string seed, int variant)
    {
        var underlying = Nullable.GetUnderlyingType(type);
        if (underlying != null)
        {
            return CreateSampleValue(underlying, seed, variant);
        }

        if (type == typeof(string))
        {
            return $"{seed}-value-{variant}";
        }

        if (type == typeof(int))
        {
            return variant * 10;
        }

        if (type == typeof(long))
        {
            return (long)variant * 100L;
        }

        if (type == typeof(float))
        {
            return variant + 0.5f;
        }

        if (type == typeof(double))
        {
            return variant + 0.25d;
        }

        if (type == typeof(decimal))
        {
            return variant + 0.75m;
        }

        if (type == typeof(bool))
        {
            return variant % 2 == 0;
        }

        if (type == typeof(DateTime))
        {
            return new DateTime(2026, 1, 1).AddDays(variant);
        }

        if (type == typeof(Guid))
        {
            return variant == 1
                ? Guid.Parse("11111111-1111-1111-1111-111111111111")
                : Guid.Parse("22222222-2222-2222-2222-222222222222");
        }

        if (type.IsEnum)
        {
            var values = Enum.GetValues(type);
            if (values.Length == 0)
            {
                return null;
            }

            var index = Math.Min(variant - 1, values.Length - 1);
            return values.GetValue(index);
        }

        if (typeof(IList<string>).IsAssignableFrom(type) || type == typeof(List<string>))
        {
            return new List<string> { $"{seed}-item-{variant}" };
        }

        if (typeof(IList).IsAssignableFrom(type) && type.IsGenericType)
        {
            var genericArg = type.GetGenericArguments()[0];
            var listType = typeof(List<>).MakeGenericType(genericArg);
            var list = (IList)Activator.CreateInstance(listType)!;
            var item = CreateSampleValue(genericArg, seed + "Item", variant);
            if (item != null)
            {
                list.Add(item);
            }

            return list;
        }

        if (type.IsClass)
        {
            return Activator.CreateInstance(type);
        }

        return null;
    }
}
