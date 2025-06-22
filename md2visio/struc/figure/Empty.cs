namespace md2visio.struc.figure
{
    using System;
    using System.Collections.Concurrent;

    public static class Empty
    {
        private static readonly ConcurrentDictionary<Type, object> emptyInstances = new();

        public static T Get<T>() where T : class
        {
            return (T) emptyInstances.GetOrAdd(typeof(T), type =>
                {
                    try
                    {
#pragma warning disable CS8603 
                        return Activator.CreateInstance(type);
#pragma warning restore CS8603 
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(
                            $"Can't create empty instance of {type.Name}. " +
                            "Please ensure that the type has a public parameterless constructor, " +
                            "or manually register the instance.", ex);
                    }
                });
        }

        private static object Get(Type type)
        {
            return emptyInstances.GetOrAdd(type, t =>
            {
                try
                {
#pragma warning disable CS8603
                    return Activator.CreateInstance(t);
#pragma warning restore CS8603
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        $"Can't create empty instance of {t.Name}. " +
                        "Please ensure that the type has a public parameterless constructor, " +
                        "or manually register the instance.", ex);
                }
            });
        }

        public static bool IsEmpty<T>(this T obj) where T : class
        {
            var objType = obj.GetType();
            return obj == Get(objType);
        }
    }
}
