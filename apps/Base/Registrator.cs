using System;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Base
{
    public class Registrator
    {
        private const string regPath = @"SOFTWARE\Classes\CLSID\";
        private const string kompasLibName = "Kompas_Library";
        private const string interopKeyName = "InprocServer32";
        private const string interopKeyValue = @"\mscoree.dll";

        private readonly string _typeKeyName;
        private readonly string _interopDllPath;

        public Registrator(Type t)
        {
            _typeKeyName = $"{regPath}{{{t.GUID}}}";
            _interopDllPath = Environment.GetFolderPath(Environment.SpecialFolder.System) + interopKeyValue;
        }

        public void RegisterKompasLib()
        {
            try
            {
                using RegistryKey typeKey = Registry.LocalMachine.OpenSubKey(_typeKeyName, true);
                if (typeKey != null)
                {
                    using RegistryKey libKey = typeKey.CreateSubKey(kompasLibName);
                    using RegistryKey interopKey = typeKey.OpenSubKey(interopKeyName, true);
                    if (interopKey != null)
                    {
                        typeKey.SetValue(null, _interopDllPath);
                    }
                    else
                    {
                        throw new Exception("Failed to get registry key " + interopKeyName);
                    }
                }
                else
                {
                    throw new Exception("Failed to get registry key " + _typeKeyName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format($"Error registering Kompas class for COM-Interop:\n{ex}"));
            }
        }
		
        public void UnregisterKompasLib()
        {
            try
            {
                using RegistryKey typeKey = Registry.LocalMachine.OpenSubKey(_typeKeyName, true);
                if (typeKey != null)
                {
                    RegistryKey libKey = typeKey.OpenSubKey(kompasLibName);
                    if (libKey != null)
                    {
                        libKey.Close();
                        typeKey.DeleteSubKey(kompasLibName);
                    }
                }
                else
                {
                    throw new Exception("Failed to get registry key " + _typeKeyName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format($"Error un-registering Kompas class for COM-Interop:\n{ex}"));
            }
        }
    }
}