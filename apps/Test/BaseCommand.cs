using System;
using System.Runtime.InteropServices;
using ComHelpers;
using JetBrains.Annotations;

namespace Test
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public abstract class BaseCommand
    {
        private readonly string _libName;
        
        protected BaseCommand([NotNull] string libName)
        {
            _libName = libName;
        }
        
        protected abstract void Action(short command, short mode, object kompasObj);
        
        // ReSharper disable once UnusedMember.Global
        [return: MarshalAs(UnmanagedType.BStr)]
        public string GetLibraryName()
        {
            return _libName;
        }

        /// <summary>
        /// Main app entry
        /// </summary>
        // ReSharper disable once UnusedMember.Global
        public void ExternalRunCommand(
            // ReSharper disable once UnusedParameter.Global
            [In] short command,
            // ReSharper disable once UnusedParameter.Global
            [In] short mode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object kompasObj)
        {
            Action(command, mode, kompasObj);
        }

        #region COM Registration
        [ComRegisterFunction]
        public static void RegisterKompasLib(Type t)
        {
            new Registrator(t).RegisterKompasLib();
        }
		
        [ComUnregisterFunction]
        public static void UnregisterKompasLib(Type t)
        {
            new Registrator(t).UnregisterKompasLib();
        }
        #endregion
    }
}