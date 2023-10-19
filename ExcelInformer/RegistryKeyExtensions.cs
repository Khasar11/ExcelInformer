using Microsoft.Win32;

namespace ExcelInformer
{
    public static class RegistryKeyExtensions
    {
        public static object GetValueOrDefault(this RegistryKey key, string name, object defaultValue)
        {
            if (key.GetValue(name) == null)
            {
                key.SetValue(name, defaultValue);
                return defaultValue;
            }
            return key.GetValue(name);
        }
    }
}
